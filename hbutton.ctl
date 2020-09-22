VERSION 5.00
Begin VB.UserControl HButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   PropertyPages   =   "hbutton.ctx":0000
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ToolboxBitmap   =   "hbutton.ctx":0044
End
Attribute VB_Name = "HButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Events
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event MouseEnter()
Event MouseExit()
Event MouseDown()
Event MouseUp(ByVal Button As Integer)

'Other types
Public Enum textPosEnum
    hb_TextOver
    hb_TextLeft
    hb_TextRight
    hb_TextAbove
    hb_TextUnder
End Enum

'Internal properties
Dim m_Box As Boolean
Dim m_AlwaysBox As Boolean
Dim m_DarkBorder As Boolean
Dim m_SmallBorder As Boolean
Dim m_HoverBorder As Boolean
Dim m_menuStyle As Boolean
Dim m_NoFoc As Boolean
Dim m_TextPos As textPosEnum
Dim m_Caption As String
Dim m_HoverColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Picture As Picture
Dim m_OverPicture As Picture
Dim m_DownPicture As Picture

'Other variables
Dim Focus As Boolean
Dim hasFocus As Boolean
Dim key As Boolean
Dim over As Boolean
Dim stat As Byte

'API Calls
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long


'Draw the border of the button
Private Sub DrawLines(c1 As Long, c2 As Long, c3 As Long, c4 As Long, box As Boolean)
    mWidth = UserControl.Width / 15 - 1
    mHeight = UserControl.Height / 15 - 1
    If Not m_SmallBorder Then
        If ((box Or Focus) And Not m_NoFoc) Or m_AlwaysBox Then
            mWidth = mWidth - 1
            mHeight = mHeight - 1
            d = 1
            UserControl.Line (0, 0)-(mWidth + 1, mHeight + 1), vbButtonText, B
        Else
            d = 0
        End If
        
        UserControl.Line (d, d)-(d, mHeight), c1                'left
        UserControl.Line (d, d)-(mWidth, d), c1                 'top
        UserControl.Line (mWidth, d)-(mWidth, mHeight + 1), c2  'right
        UserControl.Line (d, mHeight)-(mWidth, mHeight), c2     'bottom
        
        If Not (stat <= 1 And m_HoverBorder) Then
            UserControl.Line (d + 1, d + 1)-(d + 1, mHeight - 1), c3
            UserControl.Line (d + 1, d + 1)-(mWidth - 1, d + 1), c3
            UserControl.Line (mWidth - 1, d + 1)-(mWidth - 1, mHeight), c4
            UserControl.Line (d + 1, mHeight - 1)-(mWidth - 1, mHeight - 1), c4
        End If
    Else 'Small border
        If ((box Or Focus) And Not m_NoFoc) Or m_AlwaysBox Then
            mWidth = mWidth - 1
            mHeight = mHeight - 1
            d = 1
            UserControl.Line (0, 0)-(mWidth + 1, mHeight + 1), vbButtonText, B
        Else
            d = 0
        End If
        
        If Not (stat <= 1 And m_HoverBorder) Then
            UserControl.Line (d, d)-(d, mHeight), IIf(stat = 2 Or (stat = 1 And Not m_HoverBorder), c1, IIf(DarkBorder And stat = 3, c1, c3)) 'left
            UserControl.Line (d, d)-(mWidth, d), IIf(stat = 2 Or (stat = 1 And Not m_HoverBorder), c1, IIf(DarkBorder And stat = 3, c1, c3)) 'top
            UserControl.Line (mWidth, d)-(mWidth, mHeight + 1), IIf(stat = 3, c2, IIf(DarkBorder And (stat = 2 Or (stat = 1 And Not m_HoverBorder)), c2, c4))                     'right
            UserControl.Line (d, mHeight)-(mWidth, mHeight), IIf(stat = 3, c2, IIf(DarkBorder And (stat = 2 Or (stat = 1 And Not m_HoverBorder)), c2, c4))                        'bottom
        End If
    End If
End Sub

'Draw the text and the picture on the button
Private Sub DrawText(ByVal X2 As Integer)
    X = X2 + Int((UserControl.Width / 15 - UserControl.TextWidth(m_Caption)) / 2)
    Y = X2 + Int((UserControl.Height / 15 - UserControl.TextHeight(m_Caption)) / 2)
    On Error GoTo err_no_p
    If m_Picture.Height > 0 Then
        On Error GoTo err_no_op
        t = False
        If m_OverPicture.Height > 0 And stat > 1 Then t = True
res_no_op:
        On Error GoTo err_no_dp
        t2 = False
        If m_DownPicture.Height > 0 And stat > 2 Then t2 = True
res_no_dp:
        If t2 Then
            hh = m_DownPicture.Height / 26.45
            ww = m_DownPicture.Width / 26.45
        ElseIf t Then
            hh = m_OverPicture.Height / 26.45
            ww = m_OverPicture.Width / 26.45
        Else
            hh = m_Picture.Height / 26.45
            ww = m_Picture.Width / 26.45
        End If
        yy = Int((UserControl.Height / 30) - hh / 2) + IIf(stat = 3, 1, 0)
        xx = Int((UserControl.Width / 30) - ww / 2) + IIf(stat = 3, 1, 0)
        
        If m_TextPos = hb_TextLeft Then
            X = X2 + Int((UserControl.Width / 15 - UserControl.TextWidth(m_Caption) - ww - 1) / 2)
            xx = X + UserControl.TextWidth(m_Caption) + 1
        ElseIf m_TextPos = hb_TextRight Then
            xx = X2 + Int((UserControl.Width / 15 - UserControl.TextWidth(m_Caption) - ww - 1) / 2)
            X = xx + ww + 1
        ElseIf m_TextPos = hb_TextUnder Then
            yy = X2 + Int((UserControl.Height / 15 - UserControl.TextHeight(m_Caption) - hh - 1) / 2)
            Y = yy + hh + 1
        ElseIf m_TextPos = hb_TextAbove Then
            Y = Y2 + Int((UserControl.Height / 15 - UserControl.TextHeight(m_Caption) - hh - 1) / 2)
            yy = Y + UserControl.TextHeight(m_Caption) + 1
        End If
            
        If t2 Then
            UserControl.PaintPicture m_DownPicture, xx, yy
        ElseIf t Then
            UserControl.PaintPicture m_OverPicture, xx, yy
        Else
            UserControl.PaintPicture m_Picture, xx, yy
        End If
    End If
res_no_p:

    UserControl.CurrentX = X
    UserControl.CurrentY = Y
    If UserControl.Enabled Then
        UserControl.ForeColor = IIf(stat = 2 Or stat = 3, m_HoverColor, m_ForeColor)
        UserControl.Print m_Caption
        UserControl.ForeColor = m_ForeColor
    Else
        UserControl.ForeColor = m_ForeColor
        xx = UserControl.ForeColor
        UserControl.ForeColor = vb3DLight
        UserControl.CurrentX = X + 1
        UserControl.CurrentY = Y + 1
        UserControl.Print m_Caption
              
        UserControl.ForeColor = vb3DShadow
        UserControl.CurrentX = X
        UserControl.CurrentY = Y
        UserControl.Print m_Caption
        UserControl.ForeColor = xx
    End If

    Exit Sub
err_no_p:    Resume res_no_p
err_no_op:   Resume res_no_op
err_no_dp:   Resume res_no_dp
End Sub

Sub setOver(Optional Force As Boolean = False)
    If stat = 2 And Not Force Then Exit Sub
    stat = 2
    UserControl.Cls
    
    DrawText 0
    If Focus And Not m_NoFoc Then
        Dim X As RECT
        X.Top = 4
        X.Left = 4
        X.Right = UserControl.Width / 15 - 4
        X.Bottom = UserControl.Height / 15 - 4
        DrawFocusRect UserControl.hdc, X
    End If
    DrawLines vb3DHighlight, vb3DDKShadow, vb3DLight, vb3DShadow, m_Box
End Sub

Sub setOut(Optional Force As Boolean = False)
    If stat = 1 And Not Force Then Exit Sub
    stat = 1
    UserControl.Cls
    
    DrawText 0
    If Focus And Not m_NoFoc Then
        Dim X As RECT
        X.Top = 4
        X.Left = 4
        X.Right = UserControl.Width / 15 - 4
        X.Bottom = UserControl.Height / 15 - 4
        DrawFocusRect UserControl.hdc, X
    End If
    If m_HoverBorder Then
        DrawLines vb3DHighlight, IIf(DarkBorder, vb3DDKShadow, vb3DShadow), UserControl.BackColor, UserControl.BackColor, m_Box
    Else
        DrawLines vb3DHighlight, vb3DDKShadow, vb3DLight, vb3DShadow, m_Box
    End If
End Sub

Sub setDown(Optional Force As Boolean = False)
    If stat = 3 And Not Force Then Exit Sub
    stat = 3
    UserControl.Cls
    
    DrawText 1
    If Focus And Not m_NoFoc Then
        Dim X As RECT
        X.Top = 4
        X.Left = 4
        X.Right = UserControl.Width / 15 - 4
        X.Bottom = UserControl.Height / 15 - 4
        DrawFocusRect UserControl.hdc, X
    End If
    DrawLines vb3DDKShadow, vb3DHighlight, vb3DShadow, vb3DLight, m_Box
End Sub

'Events

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    If Not m_NoFoc Then
        Focus = True
        setOver
    End If
    hasFocus = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        'Access key down
        setDown
        key = True
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then SendKeys "{tab}"
    If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then SendKeys "+{tab}"
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        'Access key down
        setOver
        RaiseEvent Click
        key = False
    End If
End Sub

Private Sub UserControl_LostFocus()
    Focus = False
    hasFocus = False
    If Not over Then setOut
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SetCapture UserControl.hwnd
    
    UserControl_MouseMove Button, Shift, X, Y
    RaiseEvent MouseDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If key Then Exit Sub
    overOld = over
    over = (X > 0) And (Y > 0) And (X < UserControl.Width / 15) And (Y < UserControl.Height / 15)
    down = (Button = 1)
    
    If (over And Not overOld) And ((Not down) Or hasFocus Or m_menuStyle) Then RaiseEvent MouseEnter
    If overOld And Not over Then RaiseEvent MouseExit
    
    If (over Or Focus) And Not down Then
        SetCapture UserControl.hwnd 'Capture mouse events
        setOver
    ElseIf over And down Then
        If hasFocus Or m_menuStyle Then
            SetCapture UserControl.hwnd 'Capture mouse events
            setDown
        End If
    ElseIf Not over And down Then
        If Focus Then
            setOver
        Else
            setOut
        End If
    Else
        ReleaseCapture 'Release mouse events
        If Focus Then
            setOver
        Else
            setOut
        End If
    End If
    If Not over Then ReleaseCapture
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (hasFocus Or m_menuStyle) Then
        SetCapture UserControl.hwnd 'Capture mouse events
        RaiseEvent MouseEnter
        setOver
        Exit Sub
    End If
    UserControl_MouseMove Button, Shift, X, Y
    If over And Button = 1 Then setOver: RaiseEvent Click
    RaiseEvent MouseUp(Button)
End Sub

'Repaint the control
Private Sub UserControl_Paint()
    Select Case stat
    Case 2
        setOver True
    Case 3
        setDown True
    Case Else
        setOut True
    End Select
End Sub

'Properties get/let

Public Property Get AlwaysBox() As Boolean
     AlwaysBox = m_AlwaysBox
End Property
Public Property Let AlwaysBox(ByVal New_AlwaysBox As Boolean)
    m_AlwaysBox = New_AlwaysBox
    PropertyChanged "alwaysBox"
    UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = 0
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

Public Property Get DarkBorder() As Boolean
    DarkBorder = m_DarkBorder
End Property
Public Property Let DarkBorder(ByVal New_DarkBorder As Boolean)
    m_DarkBorder = New_DarkBorder
    PropertyChanged "darkBorder"
    UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    UserControl_Paint
End Property

Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    UserControl_Paint
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    UserControl_Paint
End Property

Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    UserControl_Paint
End Property

Public Property Get FontSize() As Single
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    UserControl_Paint
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
    UserControl_Paint
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
    UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl_Paint
End Property

Public Property Get HoverBorder() As Boolean
    HoverBorder = m_HoverBorder
End Property
Public Property Let HoverBorder(ByVal New_HoverBorder As Boolean)
    m_HoverBorder = New_HoverBorder
    PropertyChanged "hoverBorder"
    UserControl_Paint
End Property

Public Property Get HoverColor() As OLE_COLOR
    HoverColor = m_HoverColor
End Property
Public Property Let HoverColor(ByVal New_HoverColor As OLE_COLOR)
    m_HoverColor = New_HoverColor
    PropertyChanged "HoverColor"
    UserControl_Paint
End Property

Public Property Get MenuStyle() As Boolean
    MenuStyle = m_menuStyle
End Property
Public Property Let MenuStyle(ByVal new_MenuStyle As Boolean)
    m_menuStyle = new_MenuStyle
    PropertyChanged "menuStyle"
End Property

Public Property Get NoFocus() As Boolean
    NoFocus = m_NoFoc
End Property
Public Property Let NoFocus(ByVal New_Box As Boolean)
    m_NoFoc = New_Box
    PropertyChanged "NoFoc"
    UserControl_Paint
End Property

Public Property Get SmallBorder() As Boolean
    SmallBorder = m_SmallBorder
End Property
Public Property Let SmallBorder(ByVal New_SmallBorder As Boolean)
    m_SmallBorder = New_SmallBorder
    PropertyChanged "smallBorder"
    UserControl_Paint
End Property

Public Property Get TextPosition() As textPosEnum
    TextPosition = m_TextPos
End Property
Public Property Let TextPosition(ByVal New_TextPos As textPosEnum)
    m_TextPos = New_TextPos
    PropertyChanged "textPos"
    UserControl_Paint
End Property

Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    UserControl_Paint
End Property

Public Property Get OverPicture() As Picture
    Set OverPicture = m_OverPicture
End Property
Public Property Set OverPicture(ByVal New_OverPicture As Picture)
    Set m_OverPicture = New_OverPicture
    PropertyChanged "OverPicture"
    UserControl_Paint
End Property

Public Property Get DownPicture() As Picture
    Set DownPicture = m_DownPicture
End Property
Public Property Set DownPicture(ByVal New_DownPicture As Picture)
    Set m_DownPicture = New_DownPicture
    PropertyChanged "DownPicture"
    UserControl_Paint
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", "Test knop")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_HoverColor = PropBag.ReadProperty("HoverColor", &H80000012)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 1)
    UserControl.FontName = PropBag.ReadProperty("FontName", "MS sans serif")
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_NoFoc = PropBag.ReadProperty("NoFocus", False)
    m_AlwaysBox = PropBag.ReadProperty("alwaysBox", False)
    m_DarkBorder = PropBag.ReadProperty("darkBorder", False)
    m_SmallBorder = PropBag.ReadProperty("smallBorder", False)
    m_HoverBorder = PropBag.ReadProperty("hoverBorder", True)
    m_menuStyle = PropBag.ReadProperty("menuStyle", False)
    m_TextPos = PropBag.ReadProperty("textPos", hb_TextOver)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_OverPicture = PropBag.ReadProperty("OverPicture", Nothing)
    Set m_DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, "Test knop")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("HoverColor", m_HoverColor, &H80000012)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("NoFocus", m_NoFoc, False)
    Call PropBag.WriteProperty("alwaysBox", m_AlwaysBox, False)
    Call PropBag.WriteProperty("darkBorder", m_DarkBorder, False)
    Call PropBag.WriteProperty("smallBorder", m_SmallBorder, False)
    Call PropBag.WriteProperty("hoverBorder", m_HoverBorder, True)
    Call PropBag.WriteProperty("menuStyle", m_menuStyle, False)
    Call PropBag.WriteProperty("textPos", m_TextPos, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("OverPicture", m_OverPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", m_DownPicture, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Box = m_def_Box
    Set m_Picture = LoadPicture("")
End Sub

Public Sub Refresh()
    UserControl.Refresh
    Select Case stat
    Case 2
        setOver True
    Case 3
        setDown True
    Case Else
        setOut True
    End Select
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

Private Sub UserControl_Show()
    setOut
End Sub
