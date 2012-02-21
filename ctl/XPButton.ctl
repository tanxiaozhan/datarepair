VERSION 5.00
Begin VB.UserControl XPButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "XPButton.ctx":0000
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' See the Readme.html file provided

'/////// Public Events sent back to the parent container
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOnButton(OnButton As Boolean)
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event DoubleClick(Button As Integer)   ' added benefit
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

' GDI32 Function Calls
' =====================================================================
' DC manipulation
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
' Shape Forming functions
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_DIFF = 4
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
' Other drawing functions
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' KERNEL32 Function Calls
' =====================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

' USER32 Function Calls
' =====================================================================
' General Windows related functions
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Standard TYPE Declarations used
' =====================================================================
Private Type POINTAPI                ' general use. Typically used for cursor location
    x As Long
    y As Long
End Type
Private Type RECT                    ' used to set/ref boundaries of a rectangle
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type BITMAP                  ' used to determine if an image is a bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type ICONINFO                ' used to determine if image is an icon
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type LOGFONT               ' used to create fonts
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type

' Custom TYPE Declarations used
' =====================================================================
Private Type ButtonDCInfo    ' used to manage the drawing DC
    hDC As Long              ' the temporary DC handle
    OldBitmap As Long        ' the original bitmap of the DC
    OldPen As Long           ' the original pen of the DC
    OldBrush As Long         ' the original brush of the DC
    ClipRgn As Long          ' used for circular button borders
    OldFont As Long          ' the original font of the DC
End Type
Private Type ButtonProperties
    bCaption As String                           ' button caption
    bCaptionAlign As AlignmentConstants          ' caption alignment (3 options)
    bCaptionStyle As CaptionEffectConstants      ' raised/sunken/default
    bBackStyle As BackStyleConstants             ' style of button (8 options)
    bStatus As Integer                           ' 0=Up, 1=Focus, 2=Down, 4=Hover
    bShape As ButtonStyleConstants               ' shape of button (rect, diagonal, circle)
    bSegPts As POINTAPI                          ' left/right offsets for diagonal button
    bRect As RECT                                ' cached caption's bounding rectangle
    bShowFocus As Boolean                        ' flag to display/hide focus rectangle
    bBackHover As Long                           ' button back color when mouse hovers
    bForeHover As Long                           ' button text color when mouse hovers
    bLockHover As HoverLockConstants             ' allows/restricts hover colors same as normal button colors (4 options)
    bGradient As GradientConstants               ' 4 gradient directions
    bGradientColor As Long                       ' Gradient color to use
    bMode As ButtonModeConstants
    bValue As Boolean
End Type
Private Type ImageProperties
    Image As StdPicture                          ' button image
    Align As ImagePlacementConstants             ' image alignment (6 options)
    Size As Integer                              ' image size (5 options)
    iRect As RECT                                ' cached image's bounding rectangle
    SourceSize As POINTAPI                       ' cached source image dimensions
    Type As Long                                 ' cached source image type (bmp/ico)
End Type

' Standard CONSTANTS as Constants or Enumerators
' =====================================================================
Private Const WHITENESS = &HFF0062
Private Const CI_BITMAP = &H0
Private Const CI_ICON = &H1
Private Const WM_KEYDOWN As Long = &H100
' //////////// Custom Colors \\\\\\\\\\\\\\\\\
Private Const vbGray = 8421504
' //////////// DrawText API Constants \\\\\\\\\\\\\\
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
' ///////////////// PROJECT-WIDE VARIABLES \\\\\\\\\\\\\\
Private ButtonDC As ButtonDCInfo       ' menu DC for drawing menu items
Private myProps As ButtonProperties    ' cached button properties
Private myImage As ImageProperties     ' cached image properties
Private bNoRefresh As Boolean          ' flag to prevent drawing when multiple properties are changing
Private curBackColor As Long           ' control's back color
Private adjBackColorUp As Long         ' cached backcolor in UP state
Private adjBackColorDn As Long         ' cached backcolor in DOWN state
Private adjHoverColor As Long          ' cached hover backcolor
Private mButton As Integer             ' used to prevent right click firing a click event
Private bTimerActive As Boolean        ' indication mouse is over a button
Private cParentBC As Long              ' parent control's back color
Private cCheckBox As Long              ' cached color for checkbox
Private bKeyDown As Boolean            ' used to properly display checkboxes

' Custom CONSTANTS as Constants or Enumerators
' =====================================================================
' ////////////// Used to set/reset HDC objects \\\\\\\\\\\\\\
Private Enum ColorObjects
    cObj_Brush = 0
    cObj_Pen = 1
    cObj_Text = 2
End Enum
' ////////////// Button Properties \\\\\\\\\\\\\\\
Public Enum ImagePlacementConstants ' image alignment
    lv_LeftEdge = 0
    lv_LeftOfCaption = 1
    lv_RightEdge = 2
    lv_RightOfCaption = 3
    lv_TopCenter = 4
    lv_BottomCenter = 5
End Enum
Public Enum ImageSizeConstants      ' image sizes
    lv_16x16 = 0
    lv_24x24 = 1
    lv_32x32 = 2
    lv_Fill_Stretch = 3
    lv_Fill_ScaleUpDown = 4
End Enum
Public Enum ButtonModeConstants
    lv_CommandButton = 0
    lv_CheckBox = 1
    lv_OptionButton = 2
End Enum
Public Enum ButtonStyleConstants    ' button shapes
    lv_Rectangular = 0
    lv_LeftDiagonal = 1
    lv_RightDiagonal = 2
    lv_FullDiagonal = 3
    lv_Round3D = 4                  ' border changes gradients when clicked
    lv_Round3DFixed = 5             ' border does not change gradients
    lv_RoundFlat = 6                ' 1-pixel black border
End Enum
Public Enum HoverLockConstants      ' hover lock options
    lv_LockTextandBackColor = 0
    lv_LockTextColorOnly = 1
    lv_LockBackColorOnly = 2
    lv_NoLocks = 3
End Enum
Public Enum GradientConstants       ' gradient directions
    lv_NoGradient = 0
    lv_Left2Right = 1
    lv_Right2Left = 2
    lv_Top2Bottom = 3
    lv_Bottom2Top = 4
End Enum
Public Enum CaptionEffectConstants  ' caption styles
    lv_default = 0
    lv_Sunken = 1
    lv_Raised = 2
End Enum
Public Enum FontStyles
    lv_PlainStyle = 0
    lv_Bold = 2
    lv_Italic = 4
    lv_Underline = 8
    lv_BoldItalic = 2 Or 4
    lv_BoldUnderline = 2 Or 8
    lv_ItalicUnderline = 4 Or 8
    lv_BoldItalicUnderline = 2 Or 4 Or 8
End Enum

Public Enum BackStyleConstants      ' button styles
    lv_w95 = 0
    lv_w31 = 1
    lv_XP = 2
    lv_Java = 3
    lv_Flat = 4
    lv_Hover = 5
    lv_Netscape = 6
    lv_Macintosh = 7
End Enum
Public Property Let ButtonStyle(style As BackStyleConstants)
Attribute ButtonStyle.VB_Description = "Various operating system button styles"

' Sets the style of button to be displayed

If style < 0 Or style > 7 Then Exit Property
myProps.bBackStyle = style
CreateButtonRegion                 ' re-create the button shape
CalculateBoundingRects 0, True     ' recalculate the text/image bounding rectangles
GetGDIMetrics "BackColor"          ' cache base colors
Refresh
PropertyChanged "BackStyle"
End Property

Public Property Get ButtonStyle() As BackStyleConstants
ButtonStyle = myProps.bBackStyle
End Property
Public Property Let Mode(nMode As ButtonModeConstants)
Attribute Mode.VB_Description = "Command button, check box or option button mode"

' Sets the button function/mode

If nMode < lv_CommandButton Or nMode > lv_OptionButton Then Exit Property
If myProps.bMode = lv_OptionButton Then
    ' option buttons. Need to remove references if the Mode changed
    If nMode < lv_OptionButton Then Call ToggleOptionButtons(-1)
End If
If myProps.bMode < lv_OptionButton And nMode = lv_OptionButton Then
    Call ToggleOptionButtons(1) ' add this instance to optionbutton collection
End If
If nMode = lv_CommandButton And myProps.bMode > lv_CommandButton Then Me.value = False
myProps.bMode = nMode
Refresh
PropertyChanged "Mode"
End Property
Public Property Get Mode() As ButtonModeConstants
Mode = myProps.bMode
End Property

Public Property Let value(bValue As Boolean)

' For option button & check box modes

If myProps.bMode = lv_CommandButton And bValue = True Then
    ' TRUE values for command buttons not allowed
    If Not UserControl.Ambient.UserMode Then
        MsgBox "This property is not applicable for command button modes.", vbInformation + vbOKOnly
    End If
    Exit Property
End If
myProps.bValue = bValue
' if optionbutton now true, need to toggle the other options buttons off
If bValue And myProps.bMode = lv_OptionButton Then Call ToggleOptionButtons(0)
Refresh
PropertyChanged "Value"
End Property
Public Property Get value() As Boolean
Attribute value.VB_Description = "Applicable to only check box or option button modes: True or False"
Attribute value.VB_UserMemId = 0
value = myProps.bValue
End Property

Public Property Let caption(sCaption As String)
Attribute caption.VB_Description = "The caption of the button. Double pipe (||) is a line break."
Attribute caption.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Attribute caption.VB_UserMemId = -518
Attribute caption.VB_MemberFlags = "200"

' Sets the button caption & hot key for the control

Dim i As Integer, j As Integer
' We look from right to left. VB uses this logic & so do I
i = InStrRev(sCaption, "&")
Do While i
    If Mid$(sCaption, i, 2) = "&&" Then
        i = InStrRev(i - 1, sCaption, "&")
    Else
        j = i + 1: i = 0
    End If
Loop
' if found, we use the next character as a hot key
If j Then AccessKeys = Mid$(sCaption, j, 1)
myProps.bCaption = sCaption                         ' cache the caption
CalculateBoundingRects 0, True                      ' recalculate button text/image bounding rects
Refresh
PropertyChanged "Caption"
End Property
Public Property Get caption() As String
caption = myProps.bCaption
End Property

Public Property Let CaptionAlign(nAlign As AlignmentConstants)
Attribute CaptionAlign.VB_Description = "Horizontal alignment of caption on the button."
Attribute CaptionAlign.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Caption options: Left, Right or Center Justified

If nAlign < vbLeftJustify Or nAlign > vbCenter Then Exit Property
If myImage.Align > lv_RightOfCaption And nAlign < vbCenter And (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
    ' also prevent left/right justifying captions when image is centered in caption
    If UserControl.Ambient.UserMode = False Then
        ' if not in user mode, then explain whey it is prevented
        MsgBox "When button images are aligned top/bottom center, " & vbCrLf & "button captions can only be center aligned", vbOKOnly + vbInformation
    End If
    Exit Property
End If
myProps.bCaptionAlign = nAlign
CalculateBoundingRects 0, True                  ' recalculate text/image bounding rects
Refresh
PropertyChanged "CapAlign"
End Property
Public Property Get CaptionAlign() As AlignmentConstants
CaptionAlign = myProps.bCaptionAlign
End Property

Public Property Let CaptionStyle(nStyle As CaptionEffectConstants)
Attribute CaptionStyle.VB_Description = "Flat, Embossed or Engraved effects"

' Sets the style, raised/sunken or flat (default)

If nStyle < lv_default Or nStyle > lv_Raised Then Exit Property
myProps.bCaptionStyle = nStyle
PropertyChanged "CapStyle"
If Len(myProps.bCaption) Then
    CalculateBoundingRects 0, True
    Refresh
End If
End Property
Public Property Get CaptionStyle() As CaptionEffectConstants
CaptionStyle = myProps.bCaptionStyle
End Property

Public Property Let ButtonShape(nShape As ButtonStyleConstants)
Attribute ButtonShape.VB_Description = "Rectangular or various diagonal shapes"
Attribute ButtonShape.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Sets the button's shape (rectangular, diagonal, or circular)

If nShape < lv_Rectangular Or nShape > lv_RoundFlat Then Exit Property
myProps.bShape = nShape
If nShape = lv_FullDiagonal Then
    If myProps.bCaptionAlign <> vbCenter Then myProps.bCaptionAlign = vbCenter
End If
bNoRefresh = False
Call UserControl_Resize
myProps.bCaptionAlign = Me.CaptionAlign
PropertyChanged "Shape"
End Property
Public Property Get ButtonShape() As ButtonStyleConstants
ButtonShape = myProps.bShape
End Property

Public Property Set Picture(xPic As StdPicture)
Attribute Picture.VB_Description = "The image used to display on the button."

' Sets the button image which to display

Set myImage.Image = xPic
If myImage.Size = 0 Then myImage.Size = 16
GetGDIMetrics "Picture"
CalculateBoundingRects 0, True              ' recalculate button's text/image bounding rects
Refresh
PropertyChanged "Image"
End Property
Public Property Get Picture() As StdPicture
Set Picture = myImage.Image
End Property

Public Property Let PictureAlign(ImgAlign As ImagePlacementConstants)
Attribute PictureAlign.VB_Description = "Alignment of the button image in relation to the caption and/or button."

' Image alignment options for button (6 different positions)

If ImgAlign < lv_LeftEdge Or ImgAlign > lv_BottomCenter Then Exit Property
myImage.Align = ImgAlign
If ImgAlign = lv_BottomCenter Or ImgAlign = lv_TopCenter Then CaptionAlign = vbCenter
CalculateBoundingRects 0, True          ' recalculate button's text/image bounding rects
Refresh
PropertyChanged "ImgAlign"
End Property
Public Property Get PictureAlign() As ImagePlacementConstants
PictureAlign = myImage.Align
End Property

Public Property Let Enabled(bEnabled As Boolean)
Attribute Enabled.VB_Description = "Determines if events are fired for this button."
Attribute Enabled.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Attribute Enabled.VB_UserMemId = -514

' Enables or disables the button

If bEnabled = UserControl.Enabled Then Exit Property
UserControl.Enabled = bEnabled
If myProps.bBackStyle = 3 And myProps.bMode = lv_CommandButton Then
    ' java disabled seems to not have the lower-left/upper-right pixels transparent
    DelayDrawing True
    CreateButtonRegion
    CalculateBoundingRects 0, True
    DelayDrawing False
Else
    Refresh
End If
PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let ShowFocusRect(bShow As Boolean)
Attribute ShowFocusRect.VB_Description = "Allows or prevents a focus rectangle from being displayed. In design mode, this may always be displayed for button set as Default."
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Shows/hides the focus rectangle when button comes into focus

myProps.bShowFocus = bShow
If ((myProps.bStatus And 1) = 1) Then
    ' if currently has the focus, then we take it off
    If Ambient.UserMode Then
        myProps.bStatus = myProps.bStatus And Not 1
        Refresh
    Else
        ' however, we don't if it is the default button
        MsgBox "The focus rectangle may appear on default buttons ONLY while in design mode, " & vbCrLf & _
            "but will not appear when the form is running.", vbInformation + vbOKOnly
    End If
Else
    Refresh
End If
PropertyChanged "Focus"
End Property
Public Property Get ShowFocusRect() As Boolean
ShowFocusRect = myProps.bShowFocus
End Property

Public Property Let PictureSize(nSize As ImageSizeConstants)
Attribute PictureSize.VB_Description = "Size of button image. Last two options automatically center image."

' Sets up to 5 picture sizes

If PictureSize < lv_16x16 Or PictureSize > lv_Fill_ScaleUpDown Then Exit Property
myImage.Size = (nSize + 2) * 8      ' I just want the pixel size
CalculateBoundingRects 0, True      ' recalculate text/image bounding rects
Refresh
PropertyChanged "ImgSize"
End Property
Public Property Get PictureSize() As ImageSizeConstants
If myImage.Size = 0 Then myImage.Size = 16
' parameters are 0,1,2,3,4 & 5, but we store them as 16,24,32,40, & 44
PictureSize = Choose(myImage.Size / 8 - 1, lv_16x16, lv_24x24, lv_32x32, lv_Fill_Stretch, lv_Fill_ScaleUpDown)
End Property

Public Property Let MousePointer(nPointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Various optional mouse pointers to use when mouse is over the button"

' Sets the mouse pointer for the button

UserControl.MousePointer = nPointer
PropertyChanged "mPointer"
End Property
Public Property Get MousePointer() As MousePointerConstants
MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(nIcon As StdPicture)
Attribute MouseIcon.VB_Description = "Icon or cursor used to display when mouse is over the button. MousePointer must be set to Custom."

' Sets the mouse icon for the button, MousePointer must be vbCustom

On Error GoTo ShowPropertyError
Set UserControl.MouseIcon = nIcon
If Not nIcon Is Nothing Then
    Me.MousePointer = vbCustom
    PropertyChanged "mIcon"
End If
Exit Property
ShowPropertyError:
If Ambient.UserMode = False Then MsgBox Err.Description, vbInformation + vbOKOnly, "Select .ico Or .cur Files Only"
End Property
Public Property Get MouseIcon() As StdPicture
Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set Font(nFont As StdFont)
Attribute Font.VB_Description = "Font used to display the caption."
Attribute Font.VB_ProcData.VB_Invoke_PropertyPutRef = ";Font"

' Sets the control's font & also the logical font to use on off-screen DC

Set UserControl.Font = nFont
GetGDIMetrics "Font"
CalculateBoundingRects 0, (Me.FontStyle > -1)    ' recalculate caption's text/image bounding rects
Refresh

PropertyChanged "Font"
End Property
Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Let FontStyle(nStyle As FontStyles)

' Allows direct changes to font attributes

With UserControl.Font
    .Bold = ((nStyle And lv_Bold) = lv_Bold)
    .Italic = ((nStyle And lv_Italic) = lv_Italic)
    .Underline = ((nStyle And lv_Underline) = lv_Underline)
End With
GetGDIMetrics "Font"
CalculateBoundingRects 0, True
PropertyChanged "Font"
Refresh
End Property
Public Property Get FontStyle() As FontStyles
Attribute FontStyle.VB_Description = "Various font attributes that can be changed directly."
Dim nStyle As Integer
nStyle = nStyle Or Abs(UserControl.Font.Bold) * 2
nStyle = nStyle Or Abs(UserControl.Font.Italic) * 4
nStyle = nStyle Or Abs(UserControl.Font.Underline) * 8
FontStyle = nStyle
End Property

Public Property Let ForeColor(nColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "The color of the caption's font ."
Attribute ForeColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Sets the caption text color

If nColor = UserControl.ForeColor Then Exit Property
UserControl.ForeColor = nColor
If myProps.bLockHover = lv_LockTextandBackColor Or myProps.bLockHover = lv_LockTextColorOnly Then
    Me.HoverForeColor = UserControl.ForeColor
End If
bNoRefresh = False
RedrawButton True
PropertyChanged "cFore"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = UserControl.ForeColor
End Property

Public Property Let BackColor(nColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Button back color. See also ResetDefaultColors"
Attribute BackColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Sets the backcolor of the button

curBackColor = ConvertColor(nColor)
If myProps.bLockHover = lv_LockBackColorOnly Or myProps.bLockHover = lv_LockTextandBackColor Then
    If myProps.bGradient Then
        Me.HoverBackColor = myProps.bGradientColor
    Else
        Me.HoverBackColor = nColor
    End If
End If
GetGDIMetrics "BackColor"
Refresh
PropertyChanged "cBack"
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = curBackColor
End Property

Public Property Let GradientColor(nColor As OLE_COLOR)

' Sets the gradient color. Gradients are used this way...
' Shade from BackColor to GradientColor
' GradientMode must be set

If (myProps.bLockHover = lv_LockTextandBackColor Or _
    myProps.bLockHover = lv_LockBackColorOnly) And _
    myProps.bGradient > lv_NoGradient Then
        myProps.bBackHover = nColor
        myProps.bBackHover = Me.HoverBackColor
End If
myProps.bGradientColor = nColor
GetGDIMetrics "BackColor"
If myProps.bGradient Then Refresh
PropertyChanged "cGradient"
End Property
Public Property Get GradientColor() As OLE_COLOR
Attribute GradientColor.VB_Description = "Secondary color used for gradient shades. The BackColor property is the primary color."
Attribute GradientColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
GradientColor = myProps.bGradientColor
End Property

Public Property Let GradientMode(nOpt As GradientConstants)

' Sets the direction of gradient shading

If nOpt < lv_NoGradient Or nOpt > lv_Bottom2Top Then Exit Property
myProps.bGradient = nOpt
If myProps.bLockHover = lv_LockBackColorOnly Or myProps.bLockHover = lv_LockTextandBackColor Then
    If nOpt > lv_NoGradient Then
        myProps.bBackHover = myProps.bGradientColor
    Else
        myProps.bBackHover = curBackColor
    End If
    myProps.bBackHover = Me.HoverBackColor
    GetGDIMetrics "BackColor"
End If
Refresh
PropertyChanged "Gradient"
End Property
Public Property Get GradientMode() As GradientConstants
Attribute GradientMode.VB_Description = "Various directions to draw the gradient shading."
Attribute GradientMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
GradientMode = myProps.bGradient
End Property

Public Property Let ResetDefaultColors(nDefault As Boolean)
Attribute ResetDefaultColors.VB_Description = "Resets button's back color and text color to Window's standard. The hover properties are also reset."
Attribute ResetDefaultColors.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Resets the BackColor, ForeColor, GradientColor,
' HoverBackColor & HoverForeColor to defaults

If Ambient.UserMode Or nDefault = False Then Exit Property
DelayDrawing True
curBackColor = vbButtonFace
Me.ForeColor = vbButtonText
Me.GradientColor = vbButtonFace
Me.GradientMode = lv_NoGradient
Me.HoverColorLocks = lv_LockTextandBackColor
myProps.bGradientColor = Me.GradientColor
GetGDIMetrics "BackColor"
DelayDrawing False
PropertyChanged "cGradient"
PropertyChanged "cBack"
End Property
Public Property Get ResetDefaultColors() As Boolean
ResetDefaultColors = False
End Property

Public Property Let HoverColorLocks(nLock As HoverLockConstants)
Attribute HoverColorLocks.VB_Description = "Can ensure the hover colors match the caption and back colors. Click for more options."
Attribute HoverColorLocks.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"

' Has two purposes.
' 1. If the lock wasn't set but is now set, then setting it will
' force HoverForeColor=ForeColor & HoverBackColor=Backcolor
' If gradeints in use, then HoverBackColor=GradientColor
' 2. If the lock was already set, then changing BackColor
' will force HoverBackColor to match. If gradients are used then
' it will force HoverBackColor to match GradientColor
' It will also force HoverForeColor to match ForeColor.

' After the locks have been set, manually changing the
' HoverForeColor, HoverBackColor will adjust/remove the lock

myProps.bLockHover = nLock
If myProps.bLockHover = lv_LockTextandBackColor Or _
    myProps.bLockHover = lv_LockBackColorOnly Then
        If myProps.bGradient Then
            myProps.bBackHover = myProps.bGradientColor
        Else
            myProps.bBackHover = curBackColor
        End If
        PropertyChanged "cBHover"
End If
If myProps.bLockHover = lv_LockTextandBackColor Or _
    myProps.bLockHover = lv_LockTextColorOnly Then
        myProps.bForeHover = UserControl.ForeColor
        PropertyChanged "cFHover"
End If
myProps.bBackHover = Me.HoverBackColor
myProps.bForeHover = Me.HoverForeColor
GetGDIMetrics "BackColor"
PropertyChanged "LockHover"
End Property
Public Property Get HoverColorLocks() As HoverLockConstants
HoverColorLocks = myProps.bLockHover
End Property

Public Property Let HoverForeColor(nColor As OLE_COLOR)
Attribute HoverForeColor.VB_Description = "Color of button caption's text when mouse is hovering over it. Affects the HoverLockColors property."
Attribute HoverForeColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Changes the text color when mouse is over the button
' Changing this property will affect the type of HoverLock

If myProps.bForeHover = nColor Then Exit Property
myProps.bForeHover = nColor
PropertyChanged "cFHover"
If nColor <> UserControl.ForeColor Then
    If myProps.bLockHover = lv_LockTextandBackColor Then
        myProps.bLockHover = lv_LockBackColorOnly
    Else
        If myProps.bLockHover = lv_LockTextColorOnly Then myProps.bLockHover = lv_NoLocks
    End If
End If
myProps.bLockHover = Me.HoverColorLocks
PropertyChanged "cFHover"
End Property
Public Property Get HoverForeColor() As OLE_COLOR
HoverForeColor = myProps.bForeHover
End Property

Public Property Let HoverBackColor(nColor As OLE_COLOR)
Attribute HoverBackColor.VB_Description = "Color of button background when mouse is hovering over it. Affects the HoverLockColors property."
Attribute HoverBackColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"

' Changes the backcolor when mouse is over the button
' Changing this property will affect the type of HoverLock

If myProps.bBackHover = nColor Then Exit Property
myProps.bBackHover = nColor
If nColor <> curBackColor Then
    If myProps.bLockHover = lv_LockTextandBackColor Then
        myProps.bLockHover = lv_LockTextColorOnly
    Else
        If myProps.bLockHover = lv_LockBackColorOnly Then myProps.bLockHover = lv_NoLocks
    End If
End If
myProps.bLockHover = Me.HoverColorLocks
GetGDIMetrics "BackColor"
PropertyChanged "cBHover"
End Property
Public Property Get HoverBackColor() As OLE_COLOR
HoverBackColor = myProps.bBackHover
End Property

Public Property Get hDC() As Long

' Makes the control's hDC availabe at runtime

hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long

' Makes the control's hWnd available at runtime

hwnd = UserControl.hwnd
End Property

' //////////////////// GENERAL FUNCTIONS, PUBLIC \\\\\\\\\\\\\\\\\\\\\
Public Sub Refresh()

' Refreshes the button & can be called from any form/module

RedrawButton True
End Sub

Public Sub DelayDrawing(bDelay As Boolean)

' Used to prevent redrawing button until all properties are set.
' Should you want to set multiple properties of the control during runtime
' call this function first with a TRUE parameter. Set your button
' attributes and then call it again with a FALSE property to update the
' button.   IMPORTANT: If called with a TRUE parameter you must
' also release it with a call and a FALSE parameter

' NOTE: this function will prevent flicker when several properties
' are being changed at once during run time. It is similar to
' the BeginPaint & EndPaint API functionality
bNoRefresh = bDelay
If bDelay = False Then Refresh
End Sub

Private Sub RedrawButton(bDrawEntireButton As Boolean)
' ==================================================
' Main switchboard routine for redrawing a button
' ==================================================
If bNoRefresh = True Then Exit Sub

Dim polyPts(0 To 15) As POINTAPI, polyColors(1 To 12) As Long
Dim ActiveStatus As Integer
Select Case myProps.bBackStyle
    Case 0:  DrawButton_Win95 polyPts(), polyColors(), ActiveStatus
    Case 1: DrawButton_Win31 polyPts(), polyColors(), ActiveStatus
    Case 2: DrawButton_WinXP polyPts(), polyColors(), ActiveStatus
    Case 3: DrawButton_Java polyPts(), polyColors(), ActiveStatus
    Case 4:  DrawButton_Flat polyPts(), polyColors(), ActiveStatus
    Case 5: DrawButton_Hover polyPts(), polyColors(), ActiveStatus
    Case 6: DrawButton_Netscape polyPts(), polyColors(), ActiveStatus
    Case 7: DrawButton_Macintosh polyPts(), polyColors(), ActiveStatus
End Select
Erase polyPts()
Erase polyColors()
GetSetOffDC False    ' copy the offscreen DC onto the control
UserControl.Refresh
End Sub

Private Function ToggleOptionButtons(nMode As Integer) As Boolean

' Function tracks option buttons for each container they are placed on
' It will 1) Toggle others to false when one is set to true
'         2) Add or remove option buttons from a collection
'         3) Query option buttons to see if one is set to true

Dim i As Integer, NrCtrls As Integer
Dim myObjRef As Long, tgtObjRef As Long

NrCtrls = GetProp(CLng(Tag), "lv_OptCount")
On Error GoTo OptionToggleError

If myProps.bValue And (NrCtrls > 0 Or nMode = 1) Then
    ' called when an option button is set to True; set others to false
    Dim optControl As XPButton
    myObjRef = ObjPtr(Me)
    For i = 1 To NrCtrls
        tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
        If tgtObjRef <> myObjRef Then
            CopyMemory optControl, tgtObjRef, &H4
            optControl.value = False
            CopyMemory optControl, 0&, &H4
        End If
    Next
End If
Select Case nMode
Case 1: ' Add instance to window db
    SetProp CLng(Tag), "lv_OptCount", NrCtrls + nMode
    SetProp CLng(Tag), "lv_Obj" & NrCtrls + nMode, ObjPtr(Me)
Case -1: ' Remove instance from window db
    Dim bOffset As Boolean
    myObjRef = ObjPtr(Me)
    For i = 1 To NrCtrls
        tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
        If tgtObjRef = myObjRef Then
            bOffset = -1
        Else
            If bOffset Then SetProp CLng(Tag), "lv_Obj" & i, tgtObjRef
        End If
    Next
    RemoveProp CLng(Tag), "lv_Obj" & i - 1
    If NrCtrls = 0 Then RemoveProp CLng(Tag), "lv_OptCount"
Case 2: ' See if any option buttons have True values
    For i = 1 To NrCtrls
        tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
        CopyMemory optControl, tgtObjRef, &H4
        If optControl.value = True Then
            i = NrCtrls + 1
            ToggleOptionButtons = True
        End If
        CopyMemory optControl, 0&, &H4
    Next
End Select
Exit Function

OptionToggleError:
Debug.Print "Err in OptionToggle: " & Err.Description
End Function

Friend Sub TimerUpdate(lvTimerID As Long)

' pretty good way to determine when cursor moves outside of any shape region
' especially useful for my diagonal/round buttons since they are not your typical
' rectangular shape.

Dim mousePt As POINTAPI, cRect As RECT
GetCursorPos mousePt
If WindowFromPoint(mousePt.x, mousePt.y) <> UserControl.hwnd Then
    ' when exits button area, kill the timer
    KillTimer UserControl.hwnd, lvTimerID
    myProps.bStatus = myProps.bStatus And Not 4
    bTimerActive = False
    bNoRefresh = False
    RaiseEvent MouseOnButton(False)
    bKeyDown = False
    Refresh
End If
End Sub

Private Sub CalculateBoundingRects(adjWidth As Long, bFullRecalc As Boolean, Optional lEdge As Long)

' Routine measures and places the rectangles to draw
' the caption and image on the control. The results
' are cached so this routine doesn't need to run
' every time the button is redrawn/painted

Dim cRect As RECT, tRect As RECT, iRect As RECT
Dim imgOffset As RECT, bImgWidthAdj As Boolean, bImgHeightAdj As Boolean
Dim rEdge As Long, iEdge As Long

' calculations needed for diagonal buttons
Select Case myProps.bShape
Case lv_RightDiagonal
    rEdge = myProps.bSegPts.y + ((ScaleWidth - myProps.bSegPts.y) \ 3)
    adjWidth = rEdge
Case lv_LeftDiagonal
    lEdge = myProps.bSegPts.x - (myProps.bSegPts.x \ 3)
    rEdge = ScaleWidth
    adjWidth = ScaleWidth - lEdge
Case lv_FullDiagonal
    lEdge = myProps.bSegPts.x - (myProps.bSegPts.x \ 3)
    rEdge = myProps.bSegPts.y + ((ScaleWidth - myProps.bSegPts.y) \ 3)
    adjWidth = rEdge - lEdge
Case Else
    adjWidth = myProps.bSegPts.y
    rEdge = ScaleWidth
End Select
' return variables to drawing switchboard
If Not bFullRecalc Then Exit Sub

If (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
    ' image in use, calculations for image rectangle
    If myImage.Size < 33 Then
        Select Case myImage.Align
        Case lv_LeftEdge, lv_LeftOfCaption
            imgOffset.Left = myImage.Size
            bImgWidthAdj = True
        Case lv_RightEdge, lv_RightOfCaption
            imgOffset.Right = myImage.Size
            bImgWidthAdj = True
        Case lv_TopCenter
            imgOffset.Top = myImage.Size
            bImgHeightAdj = True
        Case lv_BottomCenter
            imgOffset.Bottom = myImage.Size
            bImgHeightAdj = True
        End Select
    End If
End If


If Len(myProps.bCaption) Then
    Dim sCaption As String  ' note: Replace$ not compatible with VB5
    sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
    ' calculate total available button width available for text
    cRect.Right = adjWidth - 12 - (myImage.Size * Abs(CInt(bImgWidthAdj)))
    cRect.Bottom = ScaleHeight - 12 - (myImage.Size * Abs(CInt(bImgHeightAdj = True And myImage.Align > lv_RightOfCaption)))
    ' calculate size of rectangle to hold that text, using multiline flag
    'DrawText ButtonDC.hDC, sCaption, Len(sCaption), cRect, DT_CALCRECT Or DT_WORDBREAK
    'Modify for Chinese
       'th = UserControl.TextHeight("gW")
       'th = UserControl.TextWidth("W")
       'tn = LenB(StrConv(elTex, vbFromUnicode))

    DrawText ButtonDC.hDC, sCaption, LenB(StrConv(sCaption, vbFromUnicode)), cRect, DT_CALCRECT Or DT_WORDBREAK
    
    If myProps.bCaptionStyle Then
        cRect.Right = cRect.Right + 2
        cRect.Bottom = cRect.Bottom + 2
    End If
End If

' now calculate the position of the text rectangle
If Len(myProps.bCaption) Then
    tRect = cRect
    Select Case myProps.bCaptionAlign
    Case vbLeftJustify
        OffsetRect tRect, imgOffset.Left + lEdge + 6 + (Abs(CInt(imgOffset.Left > 0) * 6)), 0
    Case vbRightJustify
        OffsetRect tRect, rEdge - imgOffset.Right - 6 - cRect.Right - (Abs(CInt(imgOffset.Right > 0) * 6)), 0
    Case vbCenter
        If imgOffset.Left > 0 And myImage.Align = lv_LeftOfCaption Then
            OffsetRect tRect, (adjWidth - (imgOffset.Left + cRect.Right + 6)) \ 2 + lEdge + 6 + imgOffset.Left, 0
        Else
            If imgOffset.Right > 0 And myImage.Align = lv_RightOfCaption Then
                OffsetRect tRect, (adjWidth - (imgOffset.Right + cRect.Right + 6)) \ 2 + lEdge, 0
            Else
                OffsetRect tRect, ((adjWidth - (imgOffset.Left + imgOffset.Right)) - cRect.Right) \ 2 + lEdge + imgOffset.Left, 0
            End If
        End If
    End Select
Else
    cRect.Bottom = -3
End If
If (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
    ' finalize image rectangle position
    Select Case myImage.Align
    Case lv_LeftEdge
        iRect.Left = lEdge + 6
    Case lv_LeftOfCaption
        If Len(myProps.bCaption) Then
            iRect.Left = tRect.Left - 6 - imgOffset.Left
        Else
            iRect.Left = lEdge + 6
        End If
    Case lv_RightOfCaption
        If Len(myProps.bCaption) Then
            iRect.Left = tRect.Right + 6
        Else
            iRect.Left = rEdge - 6 - imgOffset.Right
        End If
    Case lv_RightEdge
        iRect.Left = rEdge - 6 - imgOffset.Right
    Case lv_TopCenter
        iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Top)) \ 2
        OffsetRect tRect, 0, iRect.Top + 2 + imgOffset.Top
    Case lv_BottomCenter
        iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Bottom)) \ 2 + cRect.Bottom + 3
        OffsetRect tRect, 0, iRect.Top - 2 - cRect.Bottom
    End Select
    If myImage.Align < lv_TopCenter Then
        OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
        iRect.Top = (ScaleHeight - myImage.Size) \ 2
    Else
        iRect.Left = (adjWidth - myImage.Size) \ 2 + lEdge
    End If
    iRect.Right = iRect.Left + myImage.Size
    iRect.Bottom = iRect.Top + myImage.Size
Else
    OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
End If
' sanity checks
If tRect.Top < 6 Then tRect.Top = 6
If tRect.Left < 6 Then tRect.Left = 6
If tRect.Right > rEdge Then tRect.Right = rEdge
If tRect.Bottom > ScaleHeight - 6 Then tRect.Bottom = ScaleHeight - 6
myProps.bRect = tRect
If iRect.Top < 6 Then iRect.Top = 6
If iRect.Left < 6 Then iRect.Left = 6
If iRect.Right > rEdge Then iRect.Right = rEdge
If iRect.Bottom > ScaleHeight - 6 Then iRect.Bottom = ScaleHeight - 6
myImage.iRect = iRect
End Sub

Private Sub GetSetOffDC(bSet As Boolean)

' This sets up our off screen DC & pastes results onto our control.

If bSet = True Then
    If ButtonDC.hDC = 0 Then
        ButtonDC.hDC = CreateCompatibleDC(UserControl.hDC)
        SetBkMode ButtonDC.hDC, 3&
        ' by pulling these objects now, we ensure no memory leaks &
        ' changing the objects as needed can be done in 1 line of code
        ' in the SetButtonColors routine
        ButtonDC.OldBrush = SelectObject(ButtonDC.hDC, CreateSolidBrush(0&))
        ButtonDC.OldPen = SelectObject(ButtonDC.hDC, CreatePen(0&, 1&, 0&))
        GetGDIMetrics "Font"
    End If
    If ButtonDC.OldBitmap = 0 Then
        Dim hBmp As Long
        hBmp = CreateCompatibleBitmap(UserControl.hDC, ScaleWidth, ScaleHeight)
        ButtonDC.OldBitmap = SelectObject(ButtonDC.hDC, hBmp)
    End If
Else
    BitBlt UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, ButtonDC.hDC, 0, 0, vbSrcCopy
End If
End Sub

Private Sub DrawRect(m_hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
                                   ByVal X2 As Long, ByVal Y2 As Long, _
                                   tColor As Long, Optional pColor As Long = -1, _
                                   Optional PenWidth As Long = 0, Optional PenStyle As Long = 0)

' Simple routine to draw a rectangle

If pColor <> -1 Then SetButtonColors True, m_hDC, cObj_Pen, pColor, , PenWidth, , PenStyle
SetButtonColors True, m_hDC, cObj_Brush, tColor, (pColor = -1)
Call Rectangle(m_hDC, X1, Y1, X2, Y2)
End Sub


Private Sub SetButtonColors(bSet As Boolean, m_hDC As Long, TypeObject As ColorObjects, lColor As Long, _
    Optional bSamePenColor As Boolean = True, Optional PenWidth As Long = 1, _
    Optional bSwapPens As Boolean = False, Optional PenStyle As Long = 0)

' This is the basic routine that sets a DC's pen, brush or font color

' here we store the most recent "sets" so we can reset when needed
Dim tBrush As Long, tPen As Long
If bSet Then    ' changing a DC's setting
    Select Case TypeObject
    Case cObj_Brush         ' brush is being changed
        DeleteObject SelectObject(ButtonDC.hDC, CreateSolidBrush(lColor))
        If bSamePenColor Then   ' if the pen color will be the same
            DeleteObject SelectObject(ButtonDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
        End If
    Case cObj_Pen   ' pen is being changed (mostly for drawing lines)
        DeleteObject SelectObject(ButtonDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
    Case cObj_Text  ' text color is changing
        SetTextColor m_hDC, ConvertColor(lColor)
    End Select
Else            ' resetting the DC back to the way it was
    DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldBrush)
    DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldPen)
End If
End Sub

Private Function ConvertColor(tColor As Long) As Long

' Converts VB color constants to real color values

If tColor < 0 Then
    ConvertColor = GetSysColor(tColor And &HFF&)
Else
    ConvertColor = tColor
End If
End Function

Private Sub CreateButtonRegion()

' this function creates the regions for the specific type of button style

Dim rgnA As Long, rgnB As Long, rgn2Use As Long
Dim lRatio As Single, lEdge As Long, rEdge As Long, Wd As Long
Dim ptTRI(0 To 9) As POINTAPI

myProps.bSegPts.x = 0
myProps.bSegPts.y = ScaleWidth
rgnB = CreateRectRgn(0, 0, 0, 0)

If ButtonDC.ClipRgn Then
    ' this was set for round buttons
    SelectClipRgn ButtonDC.hDC, 0
    DeleteObject ButtonDC.ClipRgn
    ButtonDC.ClipRgn = 0
End If
Select Case myProps.bShape
  Case lv_Round3D, lv_Round3DFixed, lv_RoundFlat
    rgn2Use = CreateEllipticRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1)
    ButtonDC.ClipRgn = CreateEllipticRgn(2 - Abs(myProps.bShape < lv_RoundFlat), 2 - Abs(myProps.bShape < lv_RoundFlat), ScaleWidth - 1 + Abs(myProps.bShape < lv_RoundFlat), ScaleHeight - 1 + Abs(myProps.bShape < lv_RoundFlat))
  Case lv_Rectangular
    rgn2Use = CreateRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1)
    Select Case myProps.bBackStyle
        Case 1 'Windows 16-bit
            GoSub LopOffCorners1
            GoSub LopOffCorners2
        Case 2, 7
            GoSub LopOffCorners3
            GoSub LopOffCorners4
        Case 3    'Java
            If UserControl.Enabled Then
                GoSub LopOffCorners1
                GoSub LopOffCorners2
            End If
    End Select
  
  Case Else
    ' here is my trick for ensuring a sharp edge on diagonal buttons.
    ' Basically a bastardized carpenters formula for right angles
    ' (i.e., 3+4=5 < the hypoteneus). Here I want a 60 degree angle,
    ' and not a 45 degree angle. The difference is sharp or choppy.
    ' Based off of the button height, I need to figure how much of
    ' the opposite end I need to cutoff for the diagonal edge
    lRatio = (ScaleHeight + 1) / 4
    Wd = ScaleWidth
    lEdge = (4 * lRatio)
    ' here we ensure a width of at least 5 pixels wide
    Do While Wd - lEdge < 5
        Wd = Wd + 5
    Loop
    If Wd <> ScaleWidth Then
        ' resize the control if necessary
        DelayDrawing True
        UserControl.Width = ScaleX(Wd, vbPixels, Parent.ScaleMode)
        myProps.bSegPts.y = ScaleWidth
        bNoRefresh = False
    End If
    rEdge = ScaleWidth - lEdge
    ' initial dimensions of our rectangle
    ptTRI(0).x = 0: ptTRI(0).y = 0
    ptTRI(1).x = 0
    ptTRI(1).y = ScaleHeight + 1
    ptTRI(2).x = ScaleWidth + 1
    ptTRI(2).y = ScaleHeight + 1
    ptTRI(3).x = ScaleWidth + 1
    ptTRI(3).y = 0
    ' now modify the left/right side as needed
    If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_LeftDiagonal Then
        ptTRI(1).x = lEdge  ' left portion
        myProps.bSegPts.x = lEdge
    End If
    If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_RightDiagonal Then
        ptTRI(3).x = rEdge + 1        ' bottom right
        myProps.bSegPts.y = rEdge
    End If
    ' for rounded corner buttons, we'll take of the corner pixels where appropriate when the
    ' diagonal button is not a fully-segmeneted type. Diagonal edges are always sharp,
    ' never rounded.
    rgn2Use = CreatePolygonRgn(ptTRI(0), 4, 2)
    Select Case myProps.bBackStyle
    Case 1
        If myProps.bShape = lv_RightDiagonal Then GoSub LopOffCorners1
        If myProps.bShape = lv_LeftDiagonal Then GoSub LopOffCorners2
    Case 2, 7
        If myProps.bShape = lv_RightDiagonal Then GoSub LopOffCorners3
        If myProps.bShape = lv_LeftDiagonal Then GoSub LopOffCorners4
    Case 3
        If UserControl.Enabled Then
            If myProps.bShape = lv_RightDiagonal Then GoSub LopOffCorners1
            If myProps.bShape = lv_LeftDiagonal Then GoSub LopOffCorners2
        End If
    End Select
End Select
Erase ptTRI
If rgnA Then DeleteObject rgnA
If rgnB Then DeleteObject rgnB
SetWindowRgn UserControl.hwnd, rgn2Use, True
If myProps.bSegPts.y = 0 Then myProps.bSegPts.y = ScaleWidth
Exit Sub

LopOffCorners1: ' left side top/bottom corners (Java/Win3.x)
    If myProps.bBackStyle = 3 Then
        rgnA = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
    Else
        rgnA = CreateRectRgn(0, 0, 1, 1)
    End If
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
LopOffCorners2: ' right side top/bottom corners (Java/Win3.x)
    If myProps.bBackStyle = 3 Then
        rgnA = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
    Else
        rgnA = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth - 1, ScaleHeight - 1)
    End If
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
LopOffCorners3: ' left side top/bottom corners (XP/Mac)
    ptTRI(0).x = 0: ptTRI(0).y = 0
    ptTRI(1).x = 2: ptTRI(1).y = 0
    ptTRI(2).x = 0: ptTRI(2).y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).x = 0: ptTRI(0).y = ScaleHeight
    ptTRI(1).x = 3: ptTRI(1).y = ScaleHeight
    ptTRI(2).x = 0: ptTRI(2).y = ScaleHeight - 3
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
Return
LopOffCorners4: ' right side top/bottom corners (XP/Mac)
    ptTRI(0).x = ScaleWidth: ptTRI(0).y = 0
    ptTRI(1).x = ScaleWidth - 2: ptTRI(1).y = 0
    ptTRI(2).x = ScaleWidth: ptTRI(2).y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).x = ScaleWidth: ptTRI(0).y = ScaleHeight
    ptTRI(1).x = ScaleWidth - 3: ptTRI(1).y = ScaleHeight
    ptTRI(2).x = ScaleWidth: ptTRI(2).y = ScaleHeight - 3
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
Return
End Sub

Private Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, _
                                                    lBMPsource As Long, bmpRect As RECT, _
                                                    Optional lMaskColor As Long = -1, _
                                                    Optional lNewBmpCx As Long, _
                                                    Optional lNewBmpCy As Long)
Const DSna = &H220326 '0x00220326
' =====================================================================
' A pretty good transparent bitmap maker I use in several projects
' Modified here to remove stuff I wont use (i.e., Flipping/Rotating images)
' =====================================================================

    Dim lMask2Use As Long 'COLORREF
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    Dim lHDCMem As Long, lHDCscreen As Long, lHDCsrc As Long, lHDCMask As Long, lHDCcolor As Long
    Dim x As Long, y As Long, srcX As Long, srcY As Long
    Dim lRatio(0 To 1) As Single
    Dim hPalOld As Long, hPalMem As Long
    
    lHDCscreen = GetDC(0&)
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    SelectObject lHDCsrc, lBMPsource             'Select the bitmap

        srcX = lNewBmpCx                  'Get width of bitmap
        srcY = lNewBmpCy                 'Get height of bitmap
        
        If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
        If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
        
        If (destRect.Right) = 0 Then x = lNewBmpCx Else x = (destRect.Right - destRect.Left)
        If (destRect.Bottom) = 0 Then y = lNewBmpCy Else y = (destRect.Bottom - destRect.Top)
        If lNewBmpCx > x Or lNewBmpCy > y Then
            lRatio(0) = (x / lNewBmpCx)
            lRatio(1) = (y / lNewBmpCy)
            If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)
            lNewBmpCx = lRatio(0) * lNewBmpCx
            lNewBmpCy = lRatio(0) * lNewBmpCy
            Erase lRatio
        End If
    
    lMask2Use = ConvertColor(GetPixel(lHDCsrc, 0, 0))
    
    'Create some DCs & bitmaps
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, x, y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ReleaseDC 0&, lHDCscreen
    
' ====================== Start working here ======================
    
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lHDCMem, 0, True)
    RealizePalette lHDCMem
    
    BitBlt lHDCMem, 0&, 0&, x, y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
    
    
    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
    
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna

    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    
    BitBlt lHDCdest, destRect.Left, destRect.Top, x, y, lHDCMem, 0&, 0&, vbSrcCopy
    
    'Delete memory bitmaps & DCs
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc
End Sub

Private Sub DrawButtonIcon(iRect As RECT, adjWidth As Long, lEdge As Long)

' Routine will draw the button image

If (myImage.SourceSize.x + myImage.SourceSize.y) = 0 Then Exit Sub

Dim lImgCopy As Long
Dim imgWidth As Long, imgHeight As Long
Dim rcImage As RECT, dRect As RECT
Dim bIsDown As Boolean
Const MAGICROP = &HB8074A

bIsDown = (myImage.iRect.Left <> iRect.Left)
If myImage.Size < 33 Then   ' 16, 24, 32
    imgWidth = myImage.Size
    imgHeight = myImage.Size
Else
    If myImage.Size = 40 Then   ' strectch
        iRect.Left = Abs(CInt(bIsDown))
        iRect.Top = Abs(CInt(bIsDown))
        iRect.Right = ScaleWidth
        iRect.Bottom = ScaleHeight
        imgWidth = ScaleWidth + CInt(bIsDown)
        imgHeight = ScaleHeight + CInt(bIsDown)
    Else ' scale up/down
        Dim ratio(0 To 1) As Single
        ratio(0) = (adjWidth - 12) / myImage.SourceSize.x
        ratio(1) = (ScaleHeight - 12) / myImage.SourceSize.y
        If ratio(1) < ratio(0) Then ratio(0) = ratio(1)
        ratio(1) = myImage.SourceSize.y * ratio(0)
        ratio(0) = myImage.SourceSize.x * ratio(0)
        iRect.Left = (adjWidth - CLng(ratio(0))) \ 2 + lEdge + Abs(CInt(bIsDown))
        iRect.Top = (ScaleHeight - CLng(ratio(1))) \ 2 + Abs(CInt(bIsDown))
        iRect.Right = iRect.Left + CLng(ratio(0))
        iRect.Bottom = iRect.Top + CLng(ratio(1))
        imgWidth = CLng(ratio(0))
        imgHeight = CLng(ratio(1))
        Erase ratio
    End If
End If
lImgCopy = CopyImage(myImage.Image.Handle, myImage.Type, imgWidth, imgHeight, 0)
If lImgCopy = 0 Then Exit Sub
' destination rectangle for drawing on the DC
dRect = iRect

Dim hMemDC As Long
If UserControl.Enabled Then
    hMemDC = ButtonDC.hDC
Else
    Dim hBitmap As Long, hOldBitmap As Long
    Dim hOldBrush As Long
    Dim hOldBackColor As Long, hbrShadow As Long, hbrHilite As Long
    
    ' Create a temporary DC and bitmap to hold the image
    hMemDC = CreateCompatibleDC(ButtonDC.hDC)
    hBitmap = CreateCompatibleBitmap(ButtonDC.hDC, imgWidth + 1, imgHeight + 1)
    hOldBitmap = SelectObject(hMemDC, hBitmap)
    PatBlt hMemDC, 0, 0, imgWidth, imgHeight, WHITENESS
    OffsetRect dRect, -dRect.Left, -dRect.Top
End If
    
    If myImage.Type = CI_ICON Then
        ' draw icon directly onto the temporary DC
        ' for icons, we can draw directly on the destination DC
        DrawIconEx hMemDC, dRect.Left, dRect.Top, lImgCopy, 0, 0, 0, 0, &H3
    Else
        ' draw transparent bitmap onto the temporary DC
        DrawTransparentBitmap hMemDC, dRect, lImgCopy, rcImage, , CLng(imgWidth), CLng(imgHeight)
    End If
  
If UserControl.Enabled = False Then
    hOldBackColor = SetBkColor(ButtonDC.hDC, vbWhite)
    hbrHilite = CreateSolidBrush(ShadeColor(&HC0C0C0, 36, False))
    hbrShadow = CreateSolidBrush(ShadeColor(&HC0C0C0, -36, False))
    hOldBrush = SelectObject(ButtonDC.hDC, hbrHilite)
    BitBlt ButtonDC.hDC, iRect.Left - 1, iRect.Top - 1, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
    SelectObject ButtonDC.hDC, hbrShadow
    BitBlt ButtonDC.hDC, iRect.Left, iRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
  
    SetBkColor ButtonDC.hDC, hOldBackColor
    SelectObject ButtonDC.hDC, hOldBrush
    SelectObject hMemDC, hOldBitmap
    DeleteObject hbrHilite
    If hbrShadow Then DeleteObject hbrShadow
    DeleteObject hBitmap
    DeleteDC hMemDC
End If
If myImage.Type = CI_ICON Then DestroyIcon lImgCopy Else DeleteObject lImgCopy
End Sub

Private Function ShadeColor(lColor As Long, shadeOffset As Integer, lessBlue As Boolean, _
    Optional bFocusRect As Boolean, Optional bInvert As Boolean) As Long

' Basically supply a value between -255 and +255. Positive numbers make
' the passed color lighter and negative numbers make the color darker

Dim valRGB(0 To 2) As Integer, i As Integer

CalcNewColor:
valRGB(0) = (lColor And &HFF) + shadeOffset
valRGB(1) = ((lColor And &HFF00&) / 255&) + shadeOffset
If lessBlue Then
    valRGB(2) = (lColor And &HFF0000) / &HFF00&
    valRGB(2) = valRGB(2) + ((valRGB(2) * CLng(shadeOffset)) \ &HC0)
Else
    valRGB(2) = (lColor And &HFF0000) / &HFF00& + shadeOffset
End If

For i = 0 To 2
    If valRGB(i) > 255 Then valRGB(i) = 255
    If valRGB(i) < 0 Then valRGB(i) = 0
    If bInvert = True Then valRGB(i) = Abs(255 - valRGB(i))
Next
ShadeColor = valRGB(0) + 256& * valRGB(1) + 65536 * valRGB(2)
Erase valRGB

If bFocusRect = True And (ShadeColor = vbBlack Or ShadeColor = vbWhite) Then
    shadeOffset = -shadeOffset
    If shadeOffset = 0 Then shadeOffset = 64
    GoTo CalcNewColor
End If
End Function

Private Sub GetGDIMetrics(sObject As String)

' This routine caches information we don't want to keep gathering every time a button is redrawn.

Select Case sObject
Case "Font"
    ' called when font is changed or control is initialized
    Dim newFont As LOGFONT
    newFont.lfCharSet = 1
    newFont.lfFaceName = UserControl.Font.Name & Chr$(0)
    newFont.lfHeight = (UserControl.Font.Size * -20) / Screen.TwipsPerPixelY
    newFont.lfWeight = UserControl.Font.Weight
    newFont.lfItalic = Abs(CInt(UserControl.Font.Italic))
    newFont.lfStrikeOut = Abs(CInt(UserControl.Font.Strikethrough))
    newFont.lfUnderline = Abs(CInt(UserControl.Font.Underline))
    If ButtonDC.OldFont Then
        DeleteObject SelectObject(ButtonDC.hDC, CreateFontIndirect(newFont))
    Else
        ButtonDC.OldFont = SelectObject(ButtonDC.hDC, CreateFontIndirect(newFont))
    End If
Case "Picture"
    ' get key image information
    Dim bmpInfo As BITMAP, icoInfo As ICONINFO
    If Not myImage.Image Is Nothing Then
        GetGDIObject myImage.Image.Handle, LenB(bmpInfo), bmpInfo
        If bmpInfo.bmBits = 0 Then
            GetIconInfo myImage.Image.Handle, icoInfo
            If icoInfo.hbmColor <> 0 Then
                ' downside... API creates 2 bitmaps that we need to destroy since they aren't used in this
                ' routine & are not destroyed automatically. To prevent memory leak, we destroy them here
                GetGDIObject icoInfo.hbmColor, LenB(bmpInfo), bmpInfo
                DeleteObject icoInfo.hbmColor
                If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
                myImage.Type = CI_ICON        ' flag indicating image is an icon
            End If
        Else
            myImage.Type = CI_BITMAP     ' flag indicating image is a bitmap
        End If
    End If
    myImage.SourceSize.x = bmpInfo.bmWidth
    myImage.SourceSize.y = bmpInfo.bmHeight
Case "BackColor"
    adjBackColorUp = ConvertColor(curBackColor)
    adjBackColorDn = adjBackColorUp
    adjHoverColor = ConvertColor(myProps.bBackHover)
    If myProps.bBackStyle = 7 Then
        adjBackColorUp = ShadeColor(adjBackColorUp, &H1F, False)
        adjBackColorDn = ShadeColor(vbGray, -&H10, False)
        adjHoverColor = ShadeColor(adjHoverColor, &H1F, False)
        cCheckBox = ShadeColor(vbGray, &H10, True)
    ElseIf myProps.bBackStyle = 2 Then
        adjBackColorUp = ShadeColor(adjBackColorUp, &H30, True)
        adjHoverColor = ShadeColor(adjHoverColor, &H30, True)
        adjBackColorDn = ShadeColor(adjBackColorUp, -&H20, True)
        cCheckBox = ShadeColor(vbWhite, -&H20, True)
    Else
        If myProps.bBackStyle = 3 Then
            adjBackColorDn = ShadeColor(vbGray, &HC, False)
            cCheckBox = ShadeColor(adjBackColorDn, &H1F, False)
        Else
            cCheckBox = ShadeColor(vbWhite, -&H20, False)
        End If
    End If
End Select
End Sub

Private Sub DrawButtonBackground(bColor As Long, ActiveStatus As Integer, _
                Optional bGradientColor As Long = -1, Optional bHoverColor As Long = -1)
                
' Fill the button with the appropriate backcolor

Dim i As Integer, bColor2Use As Long
Dim focusOffset As Byte, isDown As Byte

focusOffset = Abs(((myProps.bStatus And 1) = 1))
isDown = Abs((myProps.bStatus And 6) = 6)
If isDown Then ActiveStatus = 2 Else ActiveStatus = focusOffset
                            
If bHoverColor < 0 Then bHoverColor = bColor
If bTimerActive And (((myProps.bMode = lv_CommandButton And isDown = 0) Or _
    (myProps.bValue = False And myProps.bMode > lv_CommandButton))) Then
    bColor2Use = bHoverColor
Else
    bColor2Use = bColor
End If

If myProps.bShape > lv_FullDiagonal Then
    SelectClipRgn ButtonDC.hDC, 0
    If bTimerActive = True Or ((myProps.bStatus And 6) = 6) Or myProps.bBackStyle <> 5 Then
        If myProps.bShape < lv_RoundFlat Then
            ' this little trick gives us a good edge to our round button
            ' Too simple really--draw over the entire button a gradient background
            ' then set a clipping region excluding the border size and draw the
            ' rest of the button. Text/images can't overlap the border this way
            If ((myProps.bStatus And 6) = 6) = False Or myProps.bShape = lv_Round3D Then
                i = myProps.bGradient
                myProps.bGradient = 3 + Abs(((myProps.bStatus And 6) = 6))
                DrawGradient ShadeColor(cParentBC, &HAC, False), ShadeColor(cParentBC, -&H20, False)
                myProps.bGradient = i
            End If
        Else
            SetButtonColors True, ButtonDC.hDC, cObj_Pen, vbBlack, , 2
            Arc ButtonDC.hDC, 0, 0, ScaleWidth, ScaleHeight, 0, 0, 0, 0
        End If
        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
    End If
End If
If myProps.bGradient And myProps.bValue = False Then
    If bTimerActive = True And ((myProps.bStatus And 6) = 6) = False And _
        (myProps.bGradientColor <> myProps.bBackHover) Then
        DrawRect ButtonDC.hDC, 0, 0, ScaleWidth, ScaleHeight, bHoverColor
    Else
        If bGradientColor < 0 Then bGradientColor = bColor
        DrawGradient bColor, bGradientColor
    End If
Else
    If myProps.bBackStyle = 2 And (UserControl.Enabled = True Or myProps.bMode > lv_CommandButton) Then
        For i = 1 To ScaleHeight - 1
            DrawRect ButtonDC.hDC, 1, i, ScaleWidth - 1, i + 1, ShadeColor(bColor2Use, -(25 / ScaleHeight) * i, True)
        Next
    Else
        DrawRect ButtonDC.hDC, 0, 0, ScaleWidth, ScaleHeight, bColor2Use
    End If
End If
End Sub

Private Sub DrawButtonBorder(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer, _
                    Optional OuterBorderStyle As Long = -1)

' This routine draws the border depending on the button style

Dim i As Integer, j As Integer, xColorRef As Integer
Dim lBorderStyle As Long, lastColor As Long
Dim polyOffset As POINTAPI

' need to run special calculations for diagonal buttons
If myProps.bShape > lv_Rectangular And myProps.bShape < lv_Round3D Then
        polyOffset.x = Abs(CInt(myProps.bShape <> lv_RightDiagonal))
        polyOffset.y = Abs(CInt(myProps.bShape <> lv_LeftDiagonal))
End If
' calculate X,Y points for all three levels of borders
polyPts(0).x = 2 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(0).y = ScaleHeight - 3
polyPts(1).x = 2 + polyOffset.x * 2: polyPts(1).y = 2
polyPts(2).x = myProps.bSegPts.y - 3 + polyOffset.y * 3: polyPts(2).y = 2
polyPts(3).x = ScaleWidth - 3 - polyOffset.y * 3: polyPts(3).y = ScaleHeight - 3
polyPts(4).x = 1 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(4).y = ScaleHeight - 3
For i = 5 To 9
    polyPts(i).x = polyPts(i - 5).x + Choose(i - 4, polyOffset.x - 1, -1 - polyOffset.x, 1 - polyOffset.y, 1 + polyOffset.y, -1, -1)
    polyPts(i).y = polyPts(i - 5).y + Choose(i - 4, 1, -1, -1, 1, 1, 1)
Next
polyPts(10).x = myProps.bSegPts.x - polyOffset.x: polyPts(10).y = ScaleHeight - 1 + polyOffset.x
polyPts(11).x = 0: polyPts(11).y = 0
polyPts(12).x = myProps.bSegPts.y - 1 + polyOffset.y: polyPts(12).y = 0
polyPts(13).x = ScaleWidth - 1: polyPts(13).y = ScaleHeight - 1 + polyOffset.y
polyPts(14).x = myProps.bSegPts.x - 1 - polyOffset.x * 2: polyPts(14).y = ScaleHeight - 1
lastColor = -1

For i = 0 To 13
    Select Case i
        Case Is < 4: xColorRef = i + 1
        Case Is > 8:
            xColorRef = i - 1   ' next line used for dashed borders
            If OuterBorderStyle > -1 Then lBorderStyle = OuterBorderStyle
        Case Else: xColorRef = i
    End Select
    If (i <> 4 And i <> 9) Then
        ' if -1 is the color, we skip that level
        If polyColors(xColorRef) > -1 Then
            ' change the pen color if needed
            If lastColor <> polyColors(xColorRef) Then SetButtonColors True, ButtonDC.hDC, cObj_Pen, polyColors(xColorRef), , , , lBorderStyle
            Polyline ButtonDC.hDC, polyPts(i), 2
            lastColor = polyColors(xColorRef)
        End If
    End If
Next
If polyOffset.y <> ScaleWidth Then
    ' tweak to ensure bottom, outer border draws correctly on diagonal buttons
    polyPts(15).x = ScaleWidth - 1: polyPts(15).y = ScaleHeight - 1
    Polyline ButtonDC.hDC, polyPts(14), 2
End If
End Sub

Private Sub DrawFocusRectangle(fColor As Long, _
        bSolid As Boolean, bOnText As Boolean, polyPts() As POINTAPI)

' Draws focus rectangles for the button style & button mode

If ((myProps.bStatus And 1) <> 1) Then Exit Sub
If myProps.bShape > lv_Rectangular And myProps.bShape < lv_Round3D Then
    ' diagonal buttons
    Dim polyOffset As POINTAPI
    If myProps.bSegPts.x Then polyOffset.x = 1
    If myProps.bSegPts.y < ScaleWidth Then polyOffset.y = 1
    polyPts(0).x = 4 + myProps.bSegPts.x - polyOffset.x * 6: polyPts(0).y = ScaleHeight - 5
    polyPts(1).x = 4 + polyOffset.x * 4: polyPts(1).y = 4
    polyPts(2).x = myProps.bSegPts.y - 5 + polyOffset.y * 4: polyPts(2).y = 4
    polyPts(3).x = ScaleWidth - 5 - polyOffset.y * 6: polyPts(3).y = ScaleHeight - 5
    polyPts(4).x = 3 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(4).y = ScaleHeight - 5
    SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
    Polyline ButtonDC.hDC, polyPts(0), 5
Else
  Dim fRect As RECT
    Dim focusOffset As Integer
    If myProps.bShape > lv_FullDiagonal Then    ' round buttons
        SelectClipRgn ButtonDC.hDC, 0
        SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
        focusOffset = Abs(CInt(myProps.bShape = lv_RoundFlat))
        Arc ButtonDC.hDC, 3 - focusOffset, 3 - focusOffset, ScaleWidth - 3 + focusOffset, ScaleHeight - 3 + focusOffset, 0, 0, 0, 0
        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
    Else
        If bOnText = True Then
            If Len(myProps.bCaption) Then fRect = myProps.bRect Else fRect = myImage.iRect
            If bSolid Then   ' for now, only used on Java buttons
                SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
                polyPts(0).x = fRect.Left - 3: polyPts(0).y = fRect.Top - 3
                polyPts(1).x = fRect.Right + 3: polyPts(1).y = fRect.Top - 3
                polyPts(2).x = fRect.Right + 3: polyPts(2).y = fRect.Bottom + 3
                polyPts(3).x = fRect.Left - 3: polyPts(3).y = fRect.Bottom + 3
                polyPts(4).x = fRect.Left - 3: polyPts(4).y = fRect.Top - 3
                Polyline ButtonDC.hDC, polyPts(0), 5
            Else            ' for now, only used on Macintosh buttons
                focusOffset = Abs(((myProps.bStatus And 6) = 6) Or myProps.bValue = True)
                OffsetRect fRect, -3 + focusOffset, -3 + focusOffset
                fRect.Right = fRect.Right + 6
                fRect.Bottom = fRect.Bottom + 6
                DrawFocusRect ButtonDC.hDC, fRect
            End If
        Else
            fRect.Top = 0
            fRect.Left = 0
            fRect.Right = myProps.bSegPts.y - (myProps.bSegPts.x + 8)
            fRect.Bottom = ScaleHeight - 8
            OffsetRect fRect, 4 + myProps.bSegPts.x, 4
            If bSolid Then ' used when option buttons/checkboxes have focus if Value=True
                polyPts(0).x = fRect.Left: polyPts(0).y = fRect.Bottom
                polyPts(1).x = fRect.Left: polyPts(1).y = fRect.Top
                polyPts(2).x = fRect.Right: polyPts(2).y = fRect.Top
                polyPts(3).x = fRect.Right: polyPts(3).y = fRect.Bottom
                polyPts(4).x = fRect.Left: polyPts(4).y = fRect.Bottom
                SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
                Polyline ButtonDC.hDC, polyPts(0), 5
            Else
                DrawFocusRect ButtonDC.hDC, fRect
            End If
        End If
    End If
End If
End Sub

Private Sub DrawCaptionIcon(bColor As Long, Optional tColorDisabled As Long = -1, _
            Optional bOffsetTextDown As Boolean = False, _
            Optional bSingleDisableColor As Boolean = False)

' Routine draws the caption & calls the DrawButtonIcon routine

Dim tRect As RECT, iRect As RECT, adjWidth As Long, lEdge As Long
Dim lColor As Long
' let's get the adjWidth property from the following function
CalculateBoundingRects adjWidth, False, lEdge

' set these rectangles & they may be adjusted a little later
tRect = myProps.bRect
iRect = myImage.iRect
' if the button is in a down position, we'll offset the image/text rects by 1
If (((myProps.bStatus And 6) = 6) Or bOffsetTextDown) And myProps.bBackStyle <> 3 Then
    OffsetRect tRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
    OffsetRect iRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
End If

DrawButtonIcon iRect, adjWidth, lEdge
If Len(myProps.bCaption) = 0 Then Exit Sub

Dim sCaption As String  ' note Replace$ not compatible with VB5
sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
' Setting text colors and offsets
If UserControl.Enabled = False Then
    If tColorDisabled > -1 And myProps.bGradient = lv_NoGradient Then
        lColor = tColorDisabled
    Else
        lColor = vbWhite
        OffsetRect tRect, 1, 1
        bSingleDisableColor = False
    End If
Else
    ' get the right forecolor to use
    If bTimerActive = True And ((myProps.bStatus And 6) = 6) = False Then
        lColor = ConvertColor(myProps.bForeHover)
    Else
        If myProps.bGradient Then
            lColor = ConvertColor(UserControl.ForeColor)
        Else
            If myProps.bBackStyle = 7 Then
                lColor = tColorDisabled
            Else
                lColor = ConvertColor(UserControl.ForeColor)
            End If
        End If
    End If
    If (myProps.bCaptionStyle And UserControl.Enabled = True) Then
        ' drawing raised/sunken caption styles
        Dim shadeOffset As Integer
        If myProps.bCaptionStyle = lv_Raised Then shadeOffset = 40 Else shadeOffset = -40
        SetButtonColors True, ButtonDC.hDC, cObj_Text, ShadeColor(bColor, shadeOffset, False)
        OffsetRect tRect, -1, 0
        DrawText ButtonDC.hDC, sCaption, LenB(StrConv(sCaption, vbFromUnicode)), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
        SetButtonColors True, ButtonDC.hDC, cObj_Text, ShadeColor(bColor, -shadeOffset, False)
        OffsetRect tRect, 2, 2
        DrawText ButtonDC.hDC, sCaption, LenB(StrConv(sCaption, vbFromUnicode)), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
        OffsetRect tRect, -1, -1
    End If
End If
SetButtonColors True, ButtonDC.hDC, cObj_Text, lColor
DrawText ButtonDC.hDC, sCaption, LenB(StrConv(sCaption, vbFromUnicode)), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
If UserControl.Enabled = False And bSingleDisableColor = False Then
    ' finish drawing the disabled caption
    SetButtonColors True, ButtonDC.hDC, cObj_Text, vbGray
    OffsetRect tRect, -1, -1
    DrawText ButtonDC.hDC, sCaption, LenB(StrConv(sCaption, vbFromUnicode)), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
End If

End Sub

Private Sub DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long)
Dim mRect As RECT
Dim i As Long, rctOffset As Integer
Dim PixelStep As Long, rIndex As Long
Dim Colors() As Long

' The gist is to draw 1 pixel rectangles of various colors to create
' the gradient effect. If the size of the rectangle is greater than a
' quarter of the screen size, we'll step it up to 2 pixel rectangles
' to speed things up a bit


On Error Resume Next
mRect.Right = ScaleWidth
mRect.Bottom = ScaleHeight
rctOffset = 1
If myProps.bGradient < 3 Then
        If (Screen.Width \ Screen.TwipsPerPixelX) \ ScaleWidth < 4 Then
            PixelStep = ScaleWidth \ 2
            rctOffset = 2
        Else
            PixelStep = ScaleWidth
        End If
Else
    If (Screen.Height \ Screen.TwipsPerPixelY) \ ScaleHeight < 4 Then
        PixelStep = ScaleHeight \ 2
        rctOffset = 2
    Else
        PixelStep = ScaleHeight
    End If
End If
ReDim Colors(0 To PixelStep - 1) As Long
LoadGradientColors Colors(), Color1, Color2
If myProps.bGradient > 2 Then mRect.Bottom = rctOffset Else mRect.Right = rctOffset
For i = 0 To PixelStep - 1
    If myProps.bGradient Mod 2 Then rIndex = i Else rIndex = PixelStep - i - 1
    DrawRect ButtonDC.hDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(rIndex)
    If myProps.bGradient > 2 Then
        OffsetRect mRect, 0, rctOffset
    Else
        OffsetRect mRect, rctOffset, 0
    End If
Next
End Sub

Private Sub LoadGradientColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
Dim i As Integer, j As Integer
Dim sBase(0 To 2) As Single
Dim xBase(0 To 2) As Long
Dim lRatio(0 To 2) As Single

' routine adds/removes colors between a range of two colors
' Used by the DrawGradient routine. A variation of the ShadeColor routine


sBase(0) = (Color1 And &HFF)
sBase(1) = (Color1 And &HFF00&) / 255&
sBase(2) = (Color1 And &HFF0000) / &HFF00&
xBase(0) = (Color2 And &HFF)
xBase(1) = (Color2 And &HFF00&) / 255&
xBase(2) = (Color2 And &HFF0000) / &HFF00&

For j = 0 To 2
    lRatio(j) = (xBase(j) - sBase(j)) / UBound(Colors)
Next
Colors(0) = Color1
For j = 1 To UBound(Colors)
    For i = 0 To 2
        sBase(i) = sBase(i) + lRatio(i)
        If sBase(i) > 255 Then sBase(i) = 255
        If sBase(i) < 0 Then sBase(i) = 0
    Next
    Colors(j) = Int(sBase(0)) + 256& * Int(sBase(1)) + 65536 * Int(sBase(2))
Next

Erase sBase
Erase xBase
Erase lRatio
End Sub


' //////////////////// USER CONTROL EVENTS  \\\\\\\\\\\\\\\\\\\\\\\\

Private Sub UserControl_AmbientChanged(PropertyName As String)

' something on the parent container changed

On Error GoTo AbortCheck
    Select Case PropertyName
    Case "DisplayAsDefault" 'changing focus
        If Ambient.DisplayAsDefault = True And myProps.bShowFocus = True Then
            myProps.bStatus = myProps.bStatus Or 1
        Else
            myProps.bStatus = myProps.bStatus And Not 1
        End If
        Refresh
    Case "BackColor"
        cParentBC = ConvertColor(Ambient.BackColor)
        If myProps.bShape > lv_FullDiagonal Or myProps.bBackStyle = 5 Then Refresh
    End Select
AbortCheck:
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

' This happens when hot key is pressed or button is default/cancel and
' Enter/Escape key is pressed. Basically, we need to fire a click event

If (KeyAscii = 13 Or KeyAscii = 27) And myProps.bMode > lv_CommandButton Then Exit Sub
If ((myProps.bStatus And 1) <> 1) And (KeyAscii <> 13 And KeyAscii <> 27) Then Refresh
' flag that needs to be set in order to fire a click event
mButton = vbLeftButton
Call UserControl_Click  ' now trigger a click event
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
' not used by me, but we'll send the event
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

' forward arrow keys as next/previous controls

Select Case KeyCode
Case vbKeyRight
    KeyCode = 0             ' simulate a tab key
    PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H27, ByVal &H4D0001
Case vbKeyDown
    KeyCode = 0
    PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H28, ByVal &H500001
Case vbKeyLeft
    KeyCode = 0             ' simulate a shift+tab key
    PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H25, ByVal &H4B0001
Case vbKeyUp
    KeyCode = 0
    PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H26, ByVal &H480001
Case vbKeySpace
    ' space key on a button is same as enter, but shows the button state changes
    If ((myProps.bStatus And 2) <> 2) Then
        bKeyDown = True
        ' we only want to do this once. Subsequent space keys will still fire
        ' a KeyDown event, but won't keep changing button state
        ' tell routines that mouse is over button & it is "down"
        myProps.bStatus = myProps.bStatus Or 4
        myProps.bStatus = myProps.bStatus Or 2
        Refresh
        ' reset the mouse hover status if needed
        If Not bTimerActive Then myProps.bStatus = myProps.bStatus And Not 4
    End If
End Select
If KeyCode Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

' Key up events.

bKeyDown = False
Select Case KeyCode
Case vbKeySpace
    ' if space bar released & button state is "down", we make button "normal"
    If ((myProps.bStatus And 2) = 2) Then
        mButton = vbLeftButton
        myProps.bStatus = myProps.bStatus And Not 2
        If myProps.bMode > lv_CommandButton Then
            RaiseEvent KeyUp(KeyCode, Shift)
            KeyCode = 0
        Else
            Refresh
        End If
        Call UserControl_Click      ' simulate a click event
    End If
Case vbKeyRight, vbKeyDown, vbKeyLeft, vbKeyUp
    KeyCode = 0
    If myProps.bMode = lv_OptionButton And myProps.bValue = False Then
        mButton = vbLeftButton
        Call UserControl_Click
    End If
Case vbKeyShift
    KeyCode = 0
End Select
If KeyCode Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

' Only allow left clicks to fire a click event

' key variable... this tells our mouse routines & the click event
' whether or not the left button is doing the clicking
mButton = Button
If Button = vbLeftButton Then
    bKeyDown = True
    myProps.bStatus = myProps.bStatus Or 2      ' simulate a "down" state
    bNoRefresh = False
    Refresh
    ' we need this in case the user clicks & drags mouse off of the control
    ' Without it, we may never get the mouse up event
    SetCapture UserControl.hwnd
End If
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

' Here we may fire 2 events: MouseMove & MouseOnButton

On Error GoTo RaiseTheMouseEvent
' if we are already over the button, simply fire the MouseMove event
If bTimerActive Then Err.Raise 5
' if we are outside of the mouse we fire the MouseMove event only.
' Note. We don't use SetCapture/ReleaseCapture, except in one special
' case, because it affects the actual control (not my button control)
' that should rightfully have the focus. However, should the mouse
' be down on this button & the use drags mouse off of the button,
' we will continue to get mouse move events & will fire them accordingly;
' and this is the only appropriate exception for using SetCapture

If x < 0 Or y < 0 Or x > ScaleWidth Or y > ScaleHeight Then Err.Raise 5

' An improvement over most other button routines out there....
' A soft timer. No timer control needed. The trick is to get the timer
' to fire back to this instance of the button control. We do that by
' setting a reference to this instance in our Window properties. When
' the timer routine (see modLvTimer) gets an event, the hWnd is passed
' along & with that, the timer routine can retrieve the property we set.
' All this allows the timer routine to positively identify this instance.

myProps.bStatus = myProps.bStatus Or 4      ' set a mouse hover state
RaiseEvent MouseOnButton(True)              ' fire this event
' The MouseOnButton event allows users to change the properties of the
' button while the mouse is over it. For instance, you can supply a different
' image/font/etc & replace it when the mouse leaves the button area

SetProp UserControl.hwnd, "lv_ClassID", ObjPtr(Me)
' the next line is used with expandability in mind. May use multiple timers in future upgrade
SetProp UserControl.hwnd, "lv_TimerID", 237
SetTimer UserControl.hwnd, 237, 50, AddressOf lv_TimerCallBack
bTimerActive = True                         ' flag used for drawing
bNoRefresh = False                         ' ensure flag is reset
If Button = vbLeftButton Then bKeyDown = True
Refresh

RaiseTheMouseEvent:
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

' The only tweak here is to trigger a fake click event if user
' double clicked on this button

bKeyDown = False
If Button = vbLeftButton Then
    ReleaseCapture
    myProps.bStatus = myProps.bStatus And Not 2     ' "normal" state
    bNoRefresh = False                              ' ensure flag is reset
    If myProps.bMode = lv_CommandButton Then Refresh
End If
RaiseEvent MouseUp(Button, Shift, x, y)            ' fire event
' key flag. Update
mButton = Button
End Sub

Private Sub UserControl_Click()

' Again, only allow left mouse button to fire click events. Keyboard
' actions may set the mButton variable to ensure event is fired

If mButton = vbLeftButton Then
    If myProps.bMode > lv_CommandButton Then
        If myProps.bValue = True And myProps.bMode = lv_OptionButton Then Exit Sub
        Me.value = Not myProps.bValue
    End If
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_DblClick()

' Typical Window buttons do not have a double click event. Each
' double click event on a typical button is registered as 2 clicks
' with 2 sets of MouseDown & MouseUp events. We simulate that too

Dim mousePt As POINTAPI
' another plus... other button routines out there may not pass the
' true X,Y coordinates when firing a fake 2nd click event
GetCursorPos mousePt
ScreenToClient UserControl.hwnd, mousePt
RaiseEvent DoubleClick(CInt(mButton))   ' added benefit/information
If mButton = vbLeftButton Then
    ' double clicked with left mouse button fire a mouse down event
    Call UserControl_MouseDown(vbLeftButton, 0, CSng(mousePt.x), CSng(mousePt.y))
    ' key variable. This flag indicates we will be sending a fake click event
    mButton = -1
Else
    ' double clicked with middle/right mouse button, send this event only
    RaiseEvent MouseDown(vbLeftButton, 0, CSng(mousePt.x), CSng(mousePt.y))
End If
End Sub

Private Sub UserControl_GotFocus()

' If no option button in the group is set to True, then the first one that
' gets the focus is set to True by default

If myProps.bMode = lv_OptionButton And myProps.bValue = False Then
    If ToggleOptionButtons(2) = False Then
        mButton = vbLeftButton
        Call UserControl_Click
    End If
End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
' not used by me, but we'll send the event
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
' not used by me, but we'll send the event
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
' not used by me, but we'll send the event
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
' not used by me, but we'll send the event
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
' not used by me, but we'll send the event
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
' not used by me, but we'll send the event
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()

' this routine typically called by Windows when another window covering
' this button is removed, or when the parent is moved/minimized/etc.

bNoRefresh = False
RedrawButton False
End Sub

Private Sub UserControl_InitProperties()

' Initial properties for a new button

Tag = UserControl.ContainerHwnd
With myProps
    .bCaption = Ambient.DisplayName
    .bCaptionAlign = vbCenter
    .bShowFocus = True
    .bForeHover = vbButtonText
    .bBackHover = vbButtonFace
End With
cParentBC = ConvertColor(Ambient.BackColor)
curBackColor = vbButtonFace         ' this will be the button's initial backcolor
GetGDIMetrics "Font"
GetGDIMetrics "BackColor"
PropertyChanged "Caption"
PropertyChanged "CapAlign"
PropertyChanged "Focus"
PropertyChanged "cFHover"
PropertyChanged "cBHover"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

' Write properties

Tag = UserControl.ContainerHwnd
cParentBC = ConvertColor(Ambient.BackColor)
DelayDrawing True
With PropBag
    myProps.bCaption = .ReadProperty("Caption", "")
    myProps.bCaptionAlign = .ReadProperty("CapAlign", 2)
    myProps.bBackStyle = .ReadProperty("BackStyle", 0)
    myProps.bShape = .ReadProperty("Shape", 0)
    myProps.bGradient = .ReadProperty("Gradient", 0)
    myProps.bGradientColor = .ReadProperty("cGradient", vbButtonFace)
    UserControl.ForeColor = .ReadProperty("cFore", vbButtonText)
    Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
    myProps.bShowFocus = .ReadProperty("Focus", True)
    myProps.bMode = .ReadProperty("Mode", 0)
    myProps.bValue = .ReadProperty("Value", False)
    Set myImage.Image = .ReadProperty("Image", Nothing)
    myImage.Size = .ReadProperty("ImgSize", 16)
    myImage.Align = .ReadProperty("ImgAlign", 0)
    myProps.bForeHover = .ReadProperty("cFHover", vbButtonText)
    UserControl.Enabled = .ReadProperty("Enabled", True)
    curBackColor = .ReadProperty("cBack", Parent.BackColor)
    myProps.bBackHover = .ReadProperty("cBHover", curBackColor)
    myProps.bLockHover = .ReadProperty("LockHover", 0)
    myProps.bCaptionStyle = .ReadProperty("CapStyle", 0)
    Set Me.MouseIcon = .ReadProperty("mIcon", Nothing)
    Me.MousePointer = .ReadProperty("mPointer", 0)
End With
GetGDIMetrics "Picture"
GetGDIMetrics "Font"
GetGDIMetrics "BackColor"
Me.caption = myProps.bCaption      ' sets the hot key if needed
If myProps.bMode = lv_OptionButton Then ToggleOptionButtons (1)
bNoRefresh = False
Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

' Store Properties

With PropBag
    .WriteProperty "Caption", myProps.bCaption, ""
    .WriteProperty "CapAlign", myProps.bCaptionAlign, 0
    .WriteProperty "BackStyle", myProps.bBackStyle, 0
    .WriteProperty "Shape", myProps.bShape, 0
    .WriteProperty "Font", UserControl.Font, Nothing
    .WriteProperty "cFore", UserControl.ForeColor, vbButtonText
    .WriteProperty "cFHover", myProps.bForeHover, vbButtonText
    .WriteProperty "cBhover", myProps.bBackHover, curBackColor
    .WriteProperty "Focus", myProps.bShowFocus, True
    .WriteProperty "LockHover", myProps.bLockHover, 0
    .WriteProperty "cGradient", myProps.bGradientColor, vbButtonFace
    .WriteProperty "Gradient", myProps.bGradient, 0
    .WriteProperty "CapStyle", myProps.bCaptionStyle, 0
    .WriteProperty "Mode", myProps.bMode
    .WriteProperty "Value", myProps.bValue
    .WriteProperty "ImgAlign", myImage.Align, 0
    .WriteProperty "Image", myImage.Image, Nothing
    .WriteProperty "ImgSize", myImage.Size, 16
    .WriteProperty "Enabled", UserControl.Enabled, True
    .WriteProperty "cBack", curBackColor
    .WriteProperty "mPointer", UserControl.MousePointer, 0
    .WriteProperty "mIcon", UserControl.MouseIcon, Nothing
End With
End Sub

Private Sub UserControl_Resize()

' since we are using a separate DC for drawing, we need to resize the
' bitmap in that DC each time the control resizes

If bNoRefresh Then Exit Sub

CreateButtonRegion
If ButtonDC.hDC Then
    DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldBitmap)
    ButtonDC.OldBitmap = 0  ' this will force a new bitmap for existing DC
End If
GetSetOffDC True
CalculateBoundingRects 0, True
Refresh
End Sub

Private Sub UserControl_Terminate()

' Button is ending, let's clean up

' should never happen that we have a timer left over; but just in case
If bTimerActive Then KillTimer UserControl.hwnd, 1
' circular buttons have a clipping region, kill that too
If ButtonDC.ClipRgn Then DeleteObject ButtonDC.ClipRgn
If ButtonDC.hDC Then
    ' get rid of left over pen & brush
    SetButtonColors False, ButtonDC.hDC, cObj_Pen, 0
    ' get rid of logical font
    DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldFont)
    ' destroy the separate Bitmap & select original back into DC
    DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldBitmap)
    ' destroy the temporary DC
    DeleteDC ButtonDC.hDC
End If
End Sub

Private Sub DrawButton_Win95(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim midShade As Long, darkShade As Long, liteShade As Long, backShade As Long
Dim lColor As Long, fRect As RECT, i As Integer

If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
    backShade = adjBackColorUp
Else
    backShade = cCheckBox
End If

DrawButtonBackground backShade, ActiveStatus, _
                    ConvertColor(myProps.bGradientColor), adjHoverColor

If myProps.bShape < lv_Round3D Then

    If (((myProps.bStatus And 6) = 6)) Then midShade = backShade Else midShade = RGB(233, 233, 233)
    darkShade = vbGray
    liteShade = vbWhite
    If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
        ActiveStatus = 3
        midShade = RGB(233, 233, 233)
        lColor = vbBlack
    Else
        If myProps.bShape > lv_Rectangular Then lColor = ShadeColor(backShade, -&H30, False, False)
    End If
    ' inner rectangle
    polyColors(1) = Choose(ActiveStatus + 1, -1, midShade, -1, -1)
    polyColors(2) = polyColors(1)
    polyColors(3) = Choose(ActiveStatus + 1, -1, darkShade, -1, -1)
    polyColors(4) = polyColors(3)
    ' middle rectangle
    polyColors(5) = Choose(ActiveStatus + 1, midShade, liteShade, -1, -1)
    polyColors(6) = polyColors(5)
    polyColors(7) = Choose(ActiveStatus + 1, darkShade, vbBlack, darkShade, midShade)
    polyColors(8) = polyColors(7)
    ' Outer Rectangle
    If myProps.bValue = True Or (myProps.bMode > lv_CommandButton And bKeyDown = True) Then
        polyColors(9) = vbBlack: polyColors(10) = vbBlack
        polyColors(11) = vbWhite: polyColors(12) = vbWhite
    Else
        If Abs((myProps.bStatus And 1) = 1) Then polyColors(9) = vbBlack Else polyColors(9) = liteShade
        polyColors(10) = polyColors(9)
        polyColors(11) = vbBlack: polyColors(12) = vbBlack
    End If
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
End If

DrawFocusRectangle lColor, myProps.bValue, False, polyPts()
If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
    lColor = vbGray
Else
    lColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
End Sub

Private Sub DrawButton_Win31(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long
Dim i As Integer, lColor As Long

If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
    backShade = adjBackColorUp
Else
    backShade = cCheckBox
End If
DrawButtonBackground backShade, ActiveStatus, _
            ConvertColor(myProps.bGradientColor), adjHoverColor

If myProps.bShape < lv_Round3D Then

    darkShade = vbGray
    liteShade = vbWhite
    If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
        ActiveStatus = 2
        lColor = vbBlack
    Else
        If myProps.bShape > lv_Rectangular Then lColor = ShadeColor(backShade, -&H30, False, False)
        If ActiveStatus < 2 Then ActiveStatus = 1
        If myProps.bShape > lv_Rectangular Then lColor = ShadeColor(backShade, -&H30, False, False)
    End If
    If ActiveStatus = 2 Then
        polyColors(1) = darkShade
        polyColors(3) = liteShade
    Else
        polyColors(1) = liteShade
        polyColors(3) = darkShade
    End If
    polyColors(2) = polyColors(1)
    polyColors(4) = polyColors(3)
    For i = 5 To 6
        polyColors(i) = polyColors(1)
        polyColors(i + 2) = polyColors(3)
    Next
    For i = 9 To 12
        polyColors(i) = vbBlack
    Next
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
End If

DrawFocusRectangle lColor, myProps.bValue, False, polyPts()
If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
    lColor = vbGray
Else
    lColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
End Sub

Private Sub DrawButton_WinXP(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long, midShade As Long
Dim i As Integer, lColor As Long
Dim cDisabled As Long, lGradientColor As Long

If myProps.bMode > lv_CommandButton Then
    If myProps.bValue Then
        lColor = cCheckBox
    Else
        lColor = adjBackColorUp
    End If
    backShade = lColor
Else

    lColor = adjBackColorUp
    If ((myProps.bStatus And 6) = 6) And myProps.bGradient = lv_NoGradient Then
        backShade = adjBackColorDn
    Else
        If UserControl.Enabled = False Then
            backShade = ShadeColor(lColor, -&H18, True)
        Else
            backShade = lColor
        End If
    End If

End If
If Not UserControl.Enabled Then cDisabled = ShadeColor(backShade, -&H68, True)
If myProps.bGradient Then
    lGradientColor = ShadeColor(ConvertColor(myProps.bGradientColor), &H30, True)
Else
    lGradientColor = backShade
End If
DrawButtonBackground backShade, ActiveStatus, lGradientColor, adjHoverColor
    
If myProps.bShape > lv_FullDiagonal Then

    If ((myProps.bStatus And 1) = 1) And bTimerActive = False Then
        DrawFocusRectangle &HEF826B, False, False, polyPts()
    Else
        If bTimerActive = True Then DrawFocusRectangle &H96E7&, False, False, polyPts()
    End If

Else
    If UserControl.Enabled Then
        If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
           If ((myProps.bStatus And 1) = 1) Then ActiveStatus = 1 Else ActiveStatus = 2
        End If
        If ((myProps.bStatus And 4) = 4) And ((myProps.bStatus And 2) <> 2) Then
             ActiveStatus = 3
        Else
            If ActiveStatus = 1 And myProps.bShowFocus = False Then ActiveStatus = 0
        End If
        liteShade = lColor
        ' inner Rectangle
        polyColors(1) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&HA, True), &HF0D1B5, ShadeColor(liteShade, -&H16, True), &H6BCBFF)
        polyColors(2) = Choose(ActiveStatus + 1, ShadeColor(lColor, &HA, True), &HF7D7BD, ShadeColor(liteShade, -&H18, True), &H8CDBFF)
        polyColors(3) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H18, True), &HF0D1B5, lColor, &H6BCBFF)
        polyColors(4) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H20, True), &HE7AE8C, ShadeColor(liteShade, &HA, True), &H31B2FF)
        ' middle Rectangle
        polyColors(5) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H5, True), &HE7AE8C, ShadeColor(liteShade, -&H20, True), &H31B2FF)
        polyColors(6) = Choose(ActiveStatus + 1, ShadeColor(lColor, &H10, True), &HFFDFBF, ShadeColor(liteShade, -&H20, True), &HA6E9FF)
        polyColors(7) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H24, True), &HE7AE8C, ShadeColor(liteShade, &H5, True), &H31B2FF)
        polyColors(8) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H30, True), &HEF826B, ShadeColor(liteShade, &H10, True), &H96E7&)
        lColor = &H733C00
    Else
        For i = 1 To 8
            polyColors(i) = -1
        Next
        lColor = ShadeColor(lColor, -&H54, True)
    End If
    For i = 9 To 12
        polyColors(i) = lColor
    Next
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    
    If myProps.bSegPts.x = 0 Then
        SetPixel ButtonDC.hDC, 1, ScaleHeight - 2, lColor
        SetPixel ButtonDC.hDC, 1, 1, lColor
    End If
    If myProps.bSegPts.y = ScaleWidth Then
        SetPixel ButtonDC.hDC, ScaleWidth - 2, ScaleHeight - 2, lColor
        SetPixel ButtonDC.hDC, ScaleWidth - 2, 1, lColor
    End If
End If

DrawCaptionIcon backShade, cDisabled, myProps.bValue = True, True

End Sub

Private Sub DrawButton_Macintosh(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long, midShade As Long
Dim lGradientColor As Long, lFocusColor As Long
Dim i As Integer, lColor As Long

backShade = adjBackColorUp
If myProps.bMode = lv_CommandButton Then
    If ((myProps.bStatus And 6) = 6) Then backShade = adjBackColorDn
    If myProps.bShape > lv_Rectangular Then lFocusColor = ShadeColor(backShade, -&H40, False)
Else
    If myProps.bValue Then
        backShade = cCheckBox
        lFocusColor = ShadeColor(vbGray, -&H20, False)
    End If
End If

If myProps.bGradient Then
    lGradientColor = ShadeColor(ConvertColor(myProps.bGradientColor), &H1F, False)
Else
    lGradientColor = backShade
End If

DrawButtonBackground backShade, ActiveStatus, lGradientColor, adjHoverColor

If myProps.bShape < lv_Round3D Then
    
    If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
        If myProps.bValue Then ActiveStatus = 2 Else ActiveStatus = 0
    End If
    midShade = ShadeColor(backShade, &H1F, True)
    darkShade = ShadeColor(backShade, -&H40, True)
    liteShade = vbWhite
    If ActiveStatus = 2 Then
        If myProps.bGradient = lv_NoGradient Then
            lColor = vbGray
        Else
            lColor = adjBackColorUp
        End If
        backShade = lColor
        liteShade = ShadeColor(lColor, -&H20, False)
        midShade = ShadeColor(lColor, -&H40, False)
        darkShade = ShadeColor(lColor, -&H10, False)
    End If
    polyColors(1) = liteShade: polyColors(2) = liteShade
    polyColors(3) = backShade: polyColors(4) = backShade
    ' middle Rectangle
    polyColors(5) = midShade: polyColors(6) = midShade
    polyColors(7) = darkShade: polyColors(8) = darkShade
    ' Outer Rectangle
    For i = 9 To 12
        polyColors(i) = vbBlack
    Next
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    
    If myProps.bSegPts.x = 0 Then
        SetPixel ButtonDC.hDC, 3, 3, liteShade
        SetPixel ButtonDC.hDC, 1, ScaleHeight - 3, backShade
        SetPixel ButtonDC.hDC, 2, 2, midShade
        SetPixel ButtonDC.hDC, 1, ScaleHeight - 2, 0
        SetPixel ButtonDC.hDC, 1, 1, 0
    End If
    If myProps.bSegPts.y = ScaleWidth Then
        SetPixel ButtonDC.hDC, ScaleWidth - 4, ScaleHeight - 4, backShade
        SetPixel ButtonDC.hDC, ScaleWidth - 3, 1, backShade
        SetPixel ButtonDC.hDC, ScaleWidth - 3, ScaleHeight - 3, darkShade
        SetPixel ButtonDC.hDC, ScaleWidth - 2, ScaleHeight - 2, 0
        SetPixel ButtonDC.hDC, ScaleWidth - 2, 1, 0
    End If
    
End If

DrawFocusRectangle lFocusColor, myProps.bValue, True, polyPts()

If ((myProps.bStatus And 6) = 6 And myProps.bMode = lv_CommandButton) Or _
    (myProps.bMode > lv_CommandButton And myProps.bValue = True) Then
    If (myProps.bValue = True And myProps.bMode > lv_CommandButton) And UserControl.Enabled = False Then
        lColor = ShadeColor(backShade, -&H20, True)
    Else
        lColor = adjBackColorUp
    End If
Else
    If (myProps.bValue = False And myProps.bMode > lv_CommandButton) And UserControl.Enabled = False Then
        lColor = vbGray
    Else
        If UserControl.Enabled Then lColor = ConvertColor(UserControl.ForeColor) Else lColor = -1
    End If
End If
If UserControl.ForeColor = myProps.bForeHover And myProps.bValue = True Then
    lFocusColor = myProps.bForeHover
    myProps.bForeHover = lColor
Else
    lFocusColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
If lFocusColor <> -1 Then myProps.bForeHover = lFocusColor
End Sub

Private Sub DrawButton_Flat(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim darkShade As Long, liteShade As Long, backShade As Long
Dim i As Integer, lColor As Long

If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
    backShade = adjBackColorUp
Else
    backShade = cCheckBox
End If
DrawButtonBackground backShade, ActiveStatus, _
                    ConvertColor(myProps.bGradientColor), adjHoverColor

If myProps.bShape < lv_Round3D Then
    darkShade = vbGray
    liteShade = vbWhite
    ' inner rectangle & outer edges
    For i = 1 To 8
        polyColors(i) = -1
    Next
    ' Outer Rectangle
    If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
        ActiveStatus = 2
        lColor = vbBlack
    Else
        If myProps.bShape > lv_Rectangular Then lColor = ShadeColor(backShade, -&H30, False, False)
        If ActiveStatus < 2 Then ActiveStatus = 1
    End If
    polyColors(9) = Choose(ActiveStatus, liteShade, darkShade)
    polyColors(10) = polyColors(9)
    polyColors(11) = Choose(ActiveStatus, darkShade, liteShade)
    polyColors(12) = polyColors(11)
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
End If
DrawFocusRectangle lColor, myProps.bValue, False, polyPts()
If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
    lColor = vbGray
Else
    lColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
End Sub

Private Sub DrawButton_Hover(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long
Dim i As Integer, lColor As Long, lFocusColor As Long

If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
    backShade = adjBackColorUp
    If myProps.bShape > lv_Rectangular Then lFocusColor = ShadeColor(cParentBC, -&H20, False)
Else
    backShade = cCheckBox
End If
DrawButtonBackground backShade, ActiveStatus, _
            ConvertColor(myProps.bGradientColor), adjHoverColor

If myProps.bShape > lv_FullDiagonal Then
    If (Ambient.UserMode = False And backShade = cParentBC) Then
        SelectClipRgn ButtonDC.hDC, 0
        SetButtonColors True, ButtonDC.hDC, cObj_Pen, vbWhite, , , , 2
        Arc ButtonDC.hDC, 0, 0, ScaleWidth, ScaleHeight, 0, 0, 0, 0
        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
    End If
   
Else
    For i = 1 To 8
        polyColors(i) = -1
    Next
    If ((myProps.bStatus And 4) = 4) Or myProps.bValue = True Then
        If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then ActiveStatus = 2
        If ActiveStatus < 2 Then ActiveStatus = 1
        polyColors(9) = Choose(ActiveStatus, vbWhite, vbGray)
        polyColors(10) = polyColors(9)
        polyColors(11) = Choose(ActiveStatus, vbGray, vbWhite)
        polyColors(12) = polyColors(11)
    Else
        If Ambient.UserMode = False Then
            lColor = ShadeColor(cParentBC, -&H40, False)
            If lColor = vbBlack Then lColor = vbWhite
        Else
            lColor = -1
        End If
        For i = 9 To 12
            polyColors(i) = lColor
        Next
    End If
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus, Abs(UserControl.Ambient.UserMode = False) * 2
End If
DrawFocusRectangle lFocusColor, myProps.bValue, False, polyPts()
If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
    lColor = vbGray
Else
    lColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton

End Sub

Private Sub DrawButton_Netscape(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long
Dim i As Integer, lColor As Long

If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
    backShade = adjBackColorUp
Else
    backShade = cCheckBox
End If
DrawButtonBackground backShade, ActiveStatus, _
                    ConvertColor(myProps.bGradientColor), adjHoverColor

If myProps.bShape < lv_Round3D Then
    
    darkShade = vbGray
    liteShade = ShadeColor(&HDFDFDF, &H8, False)
    For i = 1 To 4
        polyColors(i) = -1
    Next
    If (bKeyDown = True And myProps.bMode > lv_CommandButton) Or myProps.bValue = True Then
        ActiveStatus = 1
        lColor = vbBlack
    Else
        ActiveStatus = Abs((myProps.bStatus And 6) = 6)
        If myProps.bShape > lv_Rectangular Then lColor = ShadeColor(backShade, -&H30, False, False)
    End If
    polyColors(5) = Choose(ActiveStatus + 1, liteShade, darkShade)
    polyColors(6) = polyColors(5): polyColors(9) = polyColors(5): polyColors(10) = polyColors(5)
    polyColors(7) = Choose(ActiveStatus + 1, darkShade, liteShade)
    polyColors(8) = polyColors(7): polyColors(11) = polyColors(7): polyColors(12) = polyColors(7)

    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
End If
DrawFocusRectangle lColor, myProps.bValue, False, polyPts()
If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
    lColor = vbGray
Else
    lColor = -1
End If
DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
End Sub

Private Sub DrawButton_Java(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long
Dim i As Integer, lColor As Long

backShade = adjBackColorUp
If myProps.bMode = lv_CommandButton Then
    If ((myProps.bStatus And 6) = 6) Then backShade = adjBackColorDn
Else
    If myProps.bValue Then backShade = cCheckBox
End If

DrawButtonBackground backShade, ActiveStatus, _
            ConvertColor(myProps.bGradientColor), adjHoverColor
lColor = vbGray
If myProps.bShape < lv_Round3D Then
    
    darkShade = ShadeColor(vbGray, -&H1A, False)
    liteShade = vbWhite
    If UserControl.Enabled Or myProps.bMode > lv_CommandButton Then
        For i = 1 To 4
            polyColors(i) = -1
        Next
        If myProps.bMode > lv_CommandButton Then
            If (myProps.bValue Or bKeyDown) Then ActiveStatus = 2 Else ActiveStatus = 3
        End If
        If ActiveStatus < 2 Then ActiveStatus = 1
        polyColors(5) = Choose(ActiveStatus, liteShade, backShade, liteShade)
        polyColors(6) = polyColors(5)
        polyColors(7) = darkShade: polyColors(8) = darkShade
        If Not UserControl.Enabled Then lColor = ShadeColor(vbGray, -&H10, False)
    Else
        For i = 1 To 8
            polyColors(i) = backShade
        Next
        liteShade = darkShade
        lColor = vbGray
    End If
    polyColors(9) = darkShade: polyColors(10) = darkShade
    polyColors(11) = liteShade: polyColors(12) = liteShade

    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
End If

DrawFocusRectangle &HCC9999, True, True, polyPts()
DrawCaptionIcon backShade, lColor, , True

End Sub

