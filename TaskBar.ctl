VERSION 5.00
Begin VB.UserControl TaskBar 
   Alignable       =   -1  'True
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ToolboxBitmap   =   "TaskBar.ctx":0000
   Begin VB.Timer tmrMouse 
      Interval        =   100
      Left            =   5760
      Top             =   120
   End
End
Attribute VB_Name = "TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' default offsets and widths
Private Const DEFAULT_ITEM_WIDTH As Single = 155
Private Const FIRST_OFFSET As Single = 1
Private Const STANDARD_OFFSET As Single = 3
Private Const ICON_WIDTH As Single = 18
'
' for drawing (button selection)
Private m_nIndexBeingSelected As Integer
Private m_bInsetSelected As Boolean
Private m_LastMouseOver As Integer
Private m_iLast As Integer

'elements linked with icons collection
'updated instantly on every change
Private m_maxCount As Integer
Public m_colIcons As Collection
Private m_refActive As clsIcon

' for drawing
Private m_cxBorder As Long
Private m_cyBorder As Long
Private m_NoDraw As Boolean

'and sizing
Private m_nOptimalHeight As Long
Private m_nAlign As AlignConstants
Private m_ActualHeight As Long
Private m_ActualWidth As Long

' tool tips
Private m_strOriginalTooltip As String
Private m_bTooltip As Boolean

' properties
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_RaisedBackColor As OLE_COLOR
Private m_SunkenBackColor As OLE_COLOR
Private m_SelectingBackColor As OLE_COLOR
Public Enum enmStyles
    Default = 0
    CoolBar = 1
End Enum
Private m_Style As enmStyles
Private m_CoolBarSeparator As Boolean
Private m_ShowActive As Boolean
Private m_AutoHide As Boolean
Private m_AutoHideWait As Integer
Private m_AutoHideAnimate As Boolean
Private m_AutoHideAnimateFrames As Integer

' property defaults
Private Const m_def_ForeColor = vbButtonText
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_RaisedBackColor = vbButtonFace
Private Const m_def_SunkenBackColor = vb3DHighlight
Private Const m_def_SelectingBackColor = vbButtonFace
Private Const m_def_Style = enmStyles.Default
Private Const m_def_CoolBarSeparator = True
Private Const m_def_ShowActive = True
Private Const m_def_AutoHide = False
Private Const m_def_AutoHideWait = 1200
Private Const m_def_AutoHideAnimate = False
Private Const m_def_AutoHideAnimateFrames = 50

' Events
Public Event ChildMinimize(ByVal hWnd As Long, ByVal Caption As String)
Public Event ChildMaximize(ByVal hWnd As Long, ByVal Caption As String)
Public Event ChildRestore(ByVal hWnd As Long, ByVal Caption As String)
Public Event ChildActivate(ByVal hWnd As Long, ByVal Caption As String)
Public Event ChildCreate(ByVal hWnd As Long, ByVal Caption As String)
Public Event ChildDestroy(ByVal hWnd As Long)
Public Event AutoHide()
Public Event AutoHideShow()

' all of these Raise* sub's are here so the module can raise
' events on the taskbar.
Public Sub RaiseChildCreate(ByVal hWnd As Long)
    
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildCreate(hWnd, sCaption)
    
End Sub

Public Sub RaiseChildDestroy(ByVal hWnd As Long)
    RaiseEvent ChildDestroy(hWnd)
End Sub

Public Sub RaiseChildMinimize(ByVal hWnd As Long)
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildMinimize(hWnd, sCaption)
'
End Sub

Public Sub RaiseChildMaximize(ByVal hWnd As Long)
    Dim sCaption As String
    sCaption = WindowText(hWnd)

    RaiseEvent ChildMaximize(hWnd, sCaption)
    
End Sub

Public Sub RaiseChildRestore(ByVal hWnd As Long)
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildRestore(hWnd, sCaption)
    
End Sub

Public Sub RaiseChildActivate(ByVal hWnd As Long)
    Dim sCaption As String
    sCaption = WindowText(hWnd)

    RaiseEvent ChildActivate(hWnd, sCaption)
    
End Sub

Friend Property Get hWnd() As Long
    ' this call is simple, it returns the usercontrols hWnd,
    ' so we can use it in GetParent() API calls in the module
    hWnd = UserControl.hWnd
End Property

Friend Sub OnRefresh(Optional ByVal hWndActive As Long = 0)
    ' we are about to refresh our taskbar
    ' called only from substitutedWndProc for MDI window
    UpdateIconsCollection hWndActive
    MapIconCollection
End Sub
'
Private Sub tmrMouse_Timer()
    ' this timer is used to take away the focus rect, for bars
    ' that are coolbar style, when the mouse leaves the
    ' window for the usercontrol.
    On Error Resume Next
    Dim oPoint As POINTAPI
    Dim lret As Long
    Dim hWnd As Long
    Dim i As Integer
    Dim lWait As Long
    Dim lTemp As Long
    Dim iCount As Integer
    Static bTest As Boolean
    Static iLast As Integer
    Static lStart As Long
    Static lNow As Long
    Static bFirstDone As Boolean
    
    If Not Ambient.UserMode() Then Exit Sub
    
    If m_Style = CoolBar Then
        ' get the X and Y of the mouse's current position
        lret = GetCursorPos(oPoint)
        ' get the handle of the window underneath that X and Y
        hWnd = WindowFromPoint(oPoint.x, oPoint.y)
        
        ' if its the first time through we have to
        ' set these up with defaults.
        If lStart = 0 And lNow = 0 Then
            lStart = GetTickCount()
            lNow = lStart
        End If
        ' if the last time we came through this sub
        ' the hWnd was NOT the same as the usercontrol.hWnd (btest = false)
        ' AND we havent re-drawn the last mouse'd over
        ' element (ilast = 1) then redraw it, to clear away
        ' the mouseover border
        If bTest = False And iLast = 1 Then
            iLast = 0
            InvalidateElement m_LastMouseOver
            m_iLast = -1
        End If
        
        If hWnd = UserControl.hWnd Then
            bTest = True
            iLast = 1
            ' handle AutoHide = True
            If m_AutoHide Then
                ' bring it back to size
                If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                    If UserControl.Extender.Height <> m_ActualHeight Then
                        UserControl.Extender.Height = m_ActualHeight
                        RaiseEvent AutoHideShow
                    End If
                ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                    If UserControl.Extender.Width <> m_ActualWidth Then
                        UserControl.Extender.Width = m_ActualWidth
                        RaiseEvent AutoHideShow
                    End If
                End If
                ' allow the paint to happen now
                m_NoDraw = False
            End If
        Else
            ' the gettickcount's are used to time how long they have been off the
            ' bar, so we can hold off on making the bar hide, for a couple
            ' seconds.
            If bTest = True Then
                ' get the starting time
                lStart = GetTickCount()
                bTest = False
            Else
                ' now every loop through, while we are off, get the new time
                lNow = GetTickCount()
            End If
        
            ' handle AutoHide = True
            ' if lNow is more than m_AutoHideWait AFTER lStart then hide
            If m_AutoHide And (((lNow - lStart) > m_AutoHideWait) Or bFirstDone = False) Then
                ' we shrink it down, because they moved the mouse
                ' outside of the usercontrol
                If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                    If m_AutoHideAnimate And UserControl.Extender.Height <> 75 And m_AutoHideAnimateFrames > 1 Then
                        ' all of this code handles the animation
                        lTemp = UserControl.Extender.Height / m_AutoHideAnimateFrames
                        iCount = m_AutoHideAnimateFrames + 1
                        Do Until iCount = 1
                            iCount = iCount - 1
                            lWait = GetTickCount()
                            If (lTemp * iCount) < 75 Then
                                Exit Do
                            End If
                            UserControl.Extender.Height = lTemp * iCount
                            Do Until (GetTickCount() - lWait) > 3
                                DoEvents
                            Loop
                        Loop
                    End If
                    UserControl.Extender.Height = 75
                    RaiseEvent AutoHide
                ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                    If m_AutoHideAnimate And UserControl.Extender.Width <> 75 And m_AutoHideAnimateFrames > 1 Then
                        lTemp = UserControl.Extender.Width / m_AutoHideAnimateFrames
                        iCount = m_AutoHideAnimateFrames + 1
                        Do Until iCount = 1
                            iCount = iCount - 1
                            lWait = GetTickCount()
                            If (lTemp * iCount) < 75 Then
                                Exit Do
                            End If
                            UserControl.Extender.Width = lTemp * iCount
                            Do Until (GetTickCount() - lWait) > 3
                                DoEvents
                            Loop
                        Loop
                    End If
                    UserControl.Extender.Width = 75
                    RaiseEvent AutoHide
                End If
                ' make sure we arent painting, and clear the usercontrol
                m_NoDraw = True
                ' bFirstDone is used to hide the bar initially if autohide
                ' is turned on
                bFirstDone = True
                UserControl.Cls
            End If
        
        End If
    
    End If
End Sub

Private Sub UserControl_Hide()
    ' if the usercontrol is going to hide, then we dont want to
    ' do any of the work we do. so stop it all
    If Ambient.UserMode() Then
        UnSubClassParentWnd Me
        ClearCollection
        On Error Resume Next
        UserControl.Parent.Arrange vbArrangeIcons
    End If
End Sub

Private Sub UserControl_InitProperties()
    If GetParent(Me.hWnd) = 0 Then
        ' no parent.
        Err.Raise 20000, "TaskBar", "TaskBar control may be placed on MDI froms only"
    Else
        ' have a parent. setup default property values
        If Ambient.UserMode() Then
            ghWnd = GetParent(hWnd)
        End If
        m_Style = m_def_Style
        m_ForeColor = m_def_ForeColor
        m_BackColor = m_def_BackColor
        m_RaisedBackColor = m_def_RaisedBackColor
        m_SunkenBackColor = m_def_SunkenBackColor
        m_SelectingBackColor = m_def_SelectingBackColor
        m_CoolBarSeparator = m_def_CoolBarSeparator
        m_ShowActive = m_def_ShowActive
        m_AutoHide = m_def_AutoHide
        m_AutoHideWait = m_def_AutoHideWait
        m_AutoHideAnimate = m_def_AutoHideAnimate
        m_AutoHideAnimateFrames = m_def_AutoHideAnimateFrames
        m_NoDraw = False
    End If
End Sub

Private Sub UserControl_LostFocus()
    ' if the usercontrol loses focus, then we
    ' want to make sure that nothing has the mouseover
    ' effect still drawn.
    Dim i As Integer
    For i = 0 To m_colIcons.Count
        InvalidateElement i
    Next i
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' mouse down, instant click (button behaviour later)
    If Not Ambient.UserMode() Then Exit Sub
    
    If Button = vbLeftButton Then
        ' detect which element and mark it
        ' set for refreshing (invalidate)
        ' set capture and wait for release
        m_nIndexBeingSelected = ElementFromPoint(x, y)
        m_bInsetSelected = (m_nIndexBeingSelected > 0)
        If (m_nIndexBeingSelected > 0) Then
            ' the set capture call just makes sure
            ' that we receive all the mouse events for our
            ' process (even if it is over another object)
            ' until we call the ReleaseCapture event.
            ' this way we can trap for the mouse up and
            ' mousemove after a mousedown, even if they
            ' happen over another part of the application.
            SetCapture UserControl.hWnd
            InvalidateElement m_nIndexBeingSelected
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim icn As clsIcon
    Dim i As Integer
    If Not Ambient.UserMode() Then Exit Sub
    
    
    If m_nIndexBeingSelected > 0 Then
        ' moving while item pressed can change state
        ' of a button originaly pressed
        Dim bNewStatus As Boolean
        ' is the mouse still over the pressed element?
        bNewStatus = IsPointInElement(x, y, m_nIndexBeingSelected)
        ' if the pressed element was one of the task bar buttons
        ' and the mouse is not over that element anymore, then
        ' take away the mouseover effect, or put the raised
        ' edge back so it looks like a butotn again.
        If m_bInsetSelected <> bNewStatus Then
            m_bInsetSelected = bNewStatus
            InvalidateElement m_nIndexBeingSelected
        End If
    ElseIf Button = 0 Then
        Dim nElPointed As Integer
        nElPointed = ElementFromPoint(x, y)
        If nElPointed > 0 Then
            Dim rc As RECT
            Dim rcRect As RECT
            Dim bDisp As Boolean
            bDisp = False
            ' if we are moving around, we want to change
            ' tooltip text (original one is in use)
            '
            ' the rule is that if we enter a button where text can't
            ' fit into, we change ToolTipText property
            ' we also set m_bTooltip flag on just to rember
            ' to restore original contns later
            If ItemRect(nElPointed, rc) Then
                If rc.Left + m_cxBorder < x And rc.Right - m_cxBorder > x And _
                    rc.Top + m_cyBorder <= y And rc.Bottom - m_cyBorder > y Then
                    bDisp = UserControl.TextWidth(m_colIcons(nElPointed).Title) > rc.Right - rc.Left - ICON_WIDTH - 6
                    ' if its using the coolbar style
                    ' it needs to have the mouseover effect.
                    If m_Style = CoolBar Then
                        ' we use the iLast so we dont keep re-drawing
                        ' it helps with the flicker.
                        If m_iLast <> nElPointed Then
                            For i = 0 To m_colIcons.Count
                                If nElPointed = i Then
                                    DrawEdge UserControl.hdc, rc, EDGE_RAISED, BF_RECT
                                    ' this is used for clearing
                                    ' the edge from a button, when
                                    ' we leave the usercontrol
                                    ' with the mouse (mouseout)
                                    m_LastMouseOver = nElPointed
                                    m_iLast = nElPointed
                                Else
                                    InvalidateElement i
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
            
            ' just setting the right tooltip
            If bDisp Then
                UserControl.Extender.ToolTipText = m_colIcons(nElPointed).Title
                m_bTooltip = True
            ElseIf m_bTooltip Then
                UserControl.Extender.ToolTipText = m_strOriginalTooltip
                m_bTooltip = False
            End If
        ElseIf m_bTooltip Then
            UserControl.Extender.ToolTipText = m_strOriginalTooltip
            m_bTooltip = False
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim lLen As Long  ' length of the string retrieved
    Dim i As Integer
    
    If Not Ambient.UserMode() Then Exit Sub
    
    'if in element we want - ok
    If m_nIndexBeingSelected > 0 And Button = vbLeftButton Then
        ' A user has released the mouse button while
        ' trying to push a button
        ' so we release the mouse capture.
        ReleaseCapture
        If IsPointInElement(x, y, m_nIndexBeingSelected) Then
            ' still inside a button from mouse down handler
            ' so change selection
            For i = 0 To m_colIcons.Count
                InvalidateElement i
            Next i
            ' since the button has been pushed, activate the window.
            ActivateWindow m_nIndexBeingSelected
        End If
        If m_bInsetSelected Then InvalidateElement m_nIndexBeingSelected
        m_nIndexBeingSelected = 0
        m_bInsetSelected = False
    ElseIf Button = vbRightButton Then
        ' raise the menu for the child, if needed
        On Error Resume Next
        Dim nNewActive As Integer
        ' get the element of the right mouse click.
        nNewActive = ElementFromPoint(x, y)
        If 0 < nNewActive Then
            ' show the menu
            ShowSystemMenu m_colIcons(nNewActive).hWnd
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    ' painting whole area
    
    Dim i As Integer
    Dim rcItem As RECT
    Dim icn As clsIcon
    Dim lEdgeParam As Long
    Dim hBrush As Long  ' receives handle to the blue hatched brush to use
    Dim r As RECT  ' rectangular area to fill
    Dim lret As Long  ' return value
    
    If Not Ambient.UserMode() Then
        Exit Sub
    End If
    
    If m_colIcons Is Nothing Or m_NoDraw = True Then
        ' no buttons, clear the control
        UserControl.Cls
        Exit Sub
    End If
    
    i = 0
    
    ' set the colors
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    
'    If m_colIcons.Count > 4 Then
'        UserControl.Height = UserControl.Height * 2
'    End If
    
    For Each icn In m_colIcons
        If ItemRect(i + 1, rcItem) Then
            ' three states of a push button
            If icn Is m_refActive Then
                ' down, currently selected button
                If m_ShowActive Then
                    ' they want to see the active button
                    DrawEdge UserControl.hdc, rcItem, EDGE_SUNKEN, BF_RECT
                    UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                    -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vb3DHighlight, BF
                    UserControl.FontBold = True
                Else
                    ' they dont want to see the active button
                    ' differently, draw it with the same code
                    ' that draws the inactvie buttons.
                    If m_Style = Default Then
                        ' if it is default then it looks like a raised button
                        DrawEdge UserControl.hdc, rcItem, EDGE_RAISED, BF_RECT
                    End If
                    UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                    -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vbButtonFace, BF
                    UserControl.FontBold = False
                End If
                
                ' now we fill the background color
                If m_ShowActive Then
                    hBrush = CreateSolidBrush(TranslateColor(m_SunkenBackColor))
                Else
                    hBrush = CreateSolidBrush(TranslateColor(m_RaisedBackColor))
                End If
            
            ElseIf i + 1 = m_nIndexBeingSelected And m_bInsetSelected Then
                ' being pushed at the moment
                DrawEdge UserControl.hdc, rcItem, EDGE_SUNKEN, BF_RECT
                UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vbButtonFace, BF
                UserControl.FontBold = False
            
                ' now we fill the background color
                hBrush = CreateSolidBrush(TranslateColor(m_SelectingBackColor))
            Else
                ' up, inactive
                If m_Style = Default Then
                    DrawEdge UserControl.hdc, rcItem, EDGE_RAISED, BF_RECT
                End If
                UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vbButtonFace, BF
                UserControl.FontBold = False
                
                ' now we fill the background color
                hBrush = CreateSolidBrush(TranslateColor(m_RaisedBackColor))
                
            End If
            
            ' we selected the right hBrush up above, now use it to
            ' draw the back color.
            lret = CopyRect(r, rcItem)
            r.Bottom = r.Bottom - 2
            r.Top = r.Top + 1
            r.Left = r.Left + 1
            r.Right = r.Right - 2
            lret = FillRect(UserControl.hdc, r, hBrush)  ' fill the rectangle using the brush
            lret = DeleteObject(hBrush)
            
            ' fix the rect to fit inside the new border thats
            ' been drawn
            rcItem.Left = rcItem.Left + m_cxBorder + 1
            rcItem.Top = rcItem.Top + m_cyBorder
            rcItem.Right = rcItem.Right - m_cxBorder - 1
            rcItem.Bottom = rcItem.Bottom - m_cyBorder - 1
            
            Dim nDiff As Single
            ' used to calculate the position to draw the icon
            nDiff = rcItem.Bottom - rcItem.Top
            
            ' draw the icon
            Dim nIconTop As Single
            Dim rcIcon As RECT
            ' calculate the position to draw the icon
            nIconTop = rcItem.Top + (nDiff - ICON_WIDTH) \ 2
            If icn.IconPtr <> 0 Then
                DrawIconEx UserControl.hdc, rcItem.Left, nIconTop + 2, icn.IconPtr, 16, 16, 0, 0, DI_NORMAL
            Else
                ' no icon was returned, so we cant draw that, we
                ' have to show the user something, so show them
                ' a big red X.
                
                ' calculate the size
                rcIcon.Left = rcItem.Left
                rcIcon.Top = nIconTop + 2
                rcIcon.Bottom = rcItem.Top + 16
                rcIcon.Right = rcItem.Left + 16
                
                ' draw a box
                UserControl.Line (rcIcon.Right, rcIcon.Top)-(rcIcon.Right, rcIcon.Bottom), vbBlack
                UserControl.Line (rcIcon.Left, rcIcon.Top)-(rcIcon.Left, rcIcon.Bottom), vbBlack
                UserControl.Line (rcIcon.Left, rcIcon.Bottom)-(rcIcon.Right, rcIcon.Bottom), vbBlack
                UserControl.Line (rcIcon.Left, rcIcon.Top)-(rcIcon.Right, rcIcon.Top), vbBlack
                
                ' now we fill the box with a big red X
                UserControl.Line (rcIcon.Left, rcIcon.Top)-(rcIcon.Right, rcIcon.Bottom), vbRed
                UserControl.Line (rcIcon.Right, rcIcon.Top)-(rcIcon.Left, rcIcon.Bottom), vbRed
            End If
            
            ' drawing a text with default font for a control
            rcItem.Left = rcItem.Left + ICON_WIDTH + 2
            Dim lpDrawTextParams As DRAWTEXTPARAMS
            lpDrawTextParams.iLeftMargin = 1
            lpDrawTextParams.iRightMargin = 1
            lpDrawTextParams.iTabLength = 2
            lpDrawTextParams.cbSize = 20
            
            ' calculate all the dimensions for the rect to
            ' draw the text in.
            Dim nTextH As Single
            nTextH = UserControl.TextHeight(icn.Title)
            If nTextH < nDiff Then
                nDiff = (nDiff - nTextH) \ 2
                rcItem.Bottom = rcItem.Bottom - nDiff
                rcItem.Top = rcItem.Top + nDiff
            End If
            ' draw the text
            DrawTextEx UserControl.hdc, icn.Title, Len(icn.Title), rcItem, _
            DT_LEFT + DT_VCENTER + DT_WORD_ELLIPSIS, lpDrawTextParams
        End If
        i = i + 1
        ' now if it is CoolBar style, we draw the separator
        ' line inbetween buttons
        If m_colIcons.Count > 1 And i < m_colIcons.Count Then
            If m_Style = CoolBar And m_CoolBarSeparator = True Then
                If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                    ' draw 2 lines to give it the 3d look.
                    UserControl.Line (r.Right + 3, r.Top)-(r.Right + 3, r.Bottom), vb3DShadow
                    UserControl.Line (r.Right + 4, r.Top + 1)-(r.Right + 4, r.Bottom + 1), TranslateColor(m_SunkenBackColor)
                ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                    UserControl.Line (r.Left, r.Bottom + 3)-(r.Right, r.Bottom + 3), vb3DShadow
                    UserControl.Line (r.Left, r.Bottom + 4)-(r.Right, r.Bottom + 4), TranslateColor(m_SunkenBackColor)
                End If
            End If
        End If
    Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoHideAnimateFrames = PropBag.ReadProperty("AutoHideAnimateFrames", m_def_AutoHideAnimateFrames)
    m_AutoHideAnimate = PropBag.ReadProperty("AutoHideAnimate", m_def_AutoHideAnimate)
    m_AutoHideWait = PropBag.ReadProperty("AutoHideWait", m_def_AutoHideWait)
    m_AutoHide = PropBag.ReadProperty("AutoHide", m_def_AutoHide)
    m_ShowActive = PropBag.ReadProperty("ShowActive", m_def_ShowActive)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_CoolBarSeparator = PropBag.ReadProperty("CoolBarSeparator", m_def_CoolBarSeparator)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_RaisedBackColor = PropBag.ReadProperty("RaisedBackColor", m_def_RaisedBackColor)
    m_SunkenBackColor = PropBag.ReadProperty("SunkenBackColor", m_def_SunkenBackColor)
    m_SelectingBackColor = PropBag.ReadProperty("SelectingBackColor", m_def_SelectingBackColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoHideAnimateFrames", m_AutoHideAnimateFrames, m_def_AutoHideAnimateFrames
    PropBag.WriteProperty "AutoHideAnimate", m_AutoHideAnimate, m_def_AutoHideAnimate
    PropBag.WriteProperty "AutoHideWait", m_AutoHideWait, m_def_AutoHideWait
    PropBag.WriteProperty "AutoHide", m_AutoHide, m_def_AutoHide
    PropBag.WriteProperty "CoolBarSeparator", m_CoolBarSeparator, m_def_CoolBarSeparator
    PropBag.WriteProperty "ShowActive", m_ShowActive, m_def_ShowActive
    PropBag.WriteProperty "Style", m_Style, m_def_Style
    PropBag.WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
    PropBag.WriteProperty "BackColor", m_BackColor, m_def_BackColor
    PropBag.WriteProperty "RaisedBackColor", m_RaisedBackColor, m_def_RaisedBackColor
    PropBag.WriteProperty "SunkenBackColor", m_SunkenBackColor, m_def_SunkenBackColor
    PropBag.WriteProperty "SelectingBackColor", m_SelectingBackColor, m_def_SelectingBackColor
End Sub

Private Sub UserControl_Show()
    ' showing the usercontrol, set it all up
    If Ambient.UserMode() Then
        SubClassParentWnd Me
        m_strOriginalTooltip = UserControl.Extender.ToolTipText
        OnRefresh
    End If
    ' get all the right heights and border widths
    m_cxBorder = GetSystemMetrics(SM_CXEDGE)
    m_cyBorder = GetSystemMetrics(SM_CYEDGE)
    m_nOptimalHeight = GetSystemMetrics(SM_CYCAPTION)
    m_nOptimalHeight = m_nOptimalHeight + 2 * m_cyBorder + 3
    
    ' set the correct height
    m_nAlign = UserControl.Extender.Align
    If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        UserControl.Extender.Height = ScaleY(m_nOptimalHeight, vbPixels, vbTwips)
        m_ActualHeight = UserControl.Extender.Height
    ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        m_ActualWidth = UserControl.Extender.Width
    End If
End Sub

Private Sub UpdateIconsCollection(ByVal hActive As Long)
    ' updates icons collection
    
    ' this function goes through every child window
    ' of the MDIClient window, and if it is a MDI Child
    ' window, it calls the VerifyItem function on it.
    ' VerifyItem just updates the information we have
    ' for a window, and adds any new windows.
On Error GoTo UpdateIconsErrorTrap
    
    Dim hWnd As Long
    Dim hWndStart As Long
    Dim sClassName As String  ' receives the name of the class
    Dim lLen As Long  ' length of the string retrieved
    
    InitCollection
    
    hWnd = FindWindowEx(ghWnd, 0, vbNullString, vbNullString)
    hWndStart = hWnd
    
    ' loop through all the child windows
    Do Until hWnd = 0
        If IsChild(ghWnd, hWnd) Then
            sClassName = Space(255)
            lLen = GetClassName(hWnd, sClassName, 255)
            sClassName = Trim(Left(sClassName, lLen))
            If sClassName = "ThunderFormDC" Or sClassName = "ThunderRT6FormDC" Then
                ' those 2 class names are hte names for MDI Child windows
                ' ThunderFormDC is the class if you are in the IDE
                ' and ThunderRT6Form is the class name if it is compiled.
                ' (no idea why microsoft did that)
                VerifyItem hWnd
            End If
        End If
        hWnd = FindWindowEx(ghWnd, hWnd, vbNullString, vbNullString)
        If hWnd = hWndStart Then
            Exit Do
        End If
    Loop
    
    Dim icn As clsIcon
    Dim iColIdx As Integer
    ' cleanup all the icons. removing the old/unneeded ones
    iColIdx = 1
    Set m_refActive = Nothing
    Do While iColIdx <= m_colIcons.Count
        Set icn = m_colIcons.Item(iColIdx)
        
        If icn.hWnd = hActive Then
            Set m_refActive = icn
            iColIdx = iColIdx + 1
        ElseIf Not icn.IsTaught Then
            m_colIcons.Remove iColIdx
        Else
            iColIdx = iColIdx + 1
        End If
    Loop
    
    m_maxCount = m_colIcons.Count
    Exit Sub
    
UpdateIconsErrorTrap:
    ' trace an error
    Debug.Print "Error occured in UpdateIconsCollection; code" + Str$(Err.Number) + " " + Err.Description
End Sub

Private Sub InitCollection()
    ' initializes collection object by clearing status
    ' of all collection elements
    If m_colIcons Is Nothing Then
        Set m_colIcons = New Collection
        m_maxCount = 0
        Exit Sub
    End If
    
    Dim icn As clsIcon
    m_maxCount = 0
    For Each icn In m_colIcons
        icn.ClearTouch
    Next
    m_maxCount = m_colIcons.Count
End Sub

Private Sub MapIconCollection()
    
    Dim icn As clsIcon
    Dim strState As String
    
    ' this function is to cause refreshing of proper parts
    ' of the control
    Static nLastPaintedCnt As Integer
    Static nLastPaintedAct As Integer
        
    If Not nLastPaintedCnt = m_colIcons.Count Then
        ' something added or removed - repaint all
        UserControl.Refresh
    Else
        Dim nElementIndex As Integer
        ' refresh every changed element taking
        ' special care of minimized windows
        ' (we hide them actually)
        nElementIndex = 1
        For Each icn In m_colIcons
            If icn.IsNew Then
                If icn.State = vbMinimized Then
                    ShowWindow icn.hWnd, SW_HIDE
                End If
                UserControl.Refresh
                Exit For
            ElseIf icn.IsChanged Then
                If icn.State = vbMinimized Then
                    ShowWindow icn.hWnd, SW_HIDE
                End If
                InvalidateElement nElementIndex
            ElseIf icn Is m_refActive And nElementIndex <> nLastPaintedAct Then
                InvalidateElement nLastPaintedAct
                InvalidateElement nElementIndex
                nLastPaintedAct = nElementIndex
            End If
            
            nElementIndex = nElementIndex + 1
        Next
    End If
    nLastPaintedCnt = m_colIcons.Count
End Sub

Private Sub VerifyItem(ByVal hWnd As Long)
    ' find item, update it if found
    ' and mark as touched
    
    Dim icn As clsIcon
    Dim lWinStyle As Long
    Dim sCaption As String
    Dim nRet As Long
    Dim hIcon As Long   ' handle to the class's icon
    Dim hLocalIcon As Long
    Dim retval As Long  ' return value
    Dim oIcon As IPictureDisp ' place for the icon
    
    For Each icn In m_colIcons
        If hWnd = icn.hWnd Then
            ' get the windows caption/title
            sCaption = WindowText(hWnd)
            
            icn.Title = sCaption
            
            ' get the WindowState of the window.
            lWinStyle = GetWindowLong(hWnd, GWL_STYLE)
            If lWinStyle And WS_MAXIMIZE Then
                icn.State = vbMaximized
            ElseIf lWinStyle And WS_MINIMIZE Then
                icn.State = vbMinimized
            Else
                icn.State = vbNormal
            End If
            
            icn.Touch
            ' we found the item, so we arent adding a new one.
            ' exit the sub before the code to add an item runs.
            Exit Sub
        End If
    Next
    
    ' new element to be added
    Set icn = New clsIcon
    
    sCaption = WindowText(hWnd)
    icn.Title = sCaption
    
    lWinStyle = GetWindowLong(hWnd, GWL_STYLE)
    If lWinStyle And WS_MAXIMIZE Then
        icn.State = vbMaximized
    ElseIf lWinStyle And WS_MINIMIZE Then
        icn.State = vbMinimized
    Else
        icn.State = vbNormal
    End If
    icn.hWnd = hWnd
    ' the IconPtr
    m_colIcons.Add icn
End Sub

Private Sub InvalidateElement(ByVal nElIdx As Integer)
    ' refreshes particular element
    ' (takes any border we have drawn off, so we can redraw it)
    If nElIdx < 1 Then Exit Sub
    
    Dim nAllCnt As Integer
    nAllCnt = m_colIcons.Count
    If nElIdx > nAllCnt Then Exit Sub
    
    'now calculate position and call invalidate rect
    Dim lpRect As RECT
    If ItemRect(nElIdx, lpRect) Then
        InvalidateRect UserControl.hWnd, lpRect, False
    End If
End Sub

Private Function ItemRect(ByVal itmIdx As Integer, ByRef rItem As RECT) As Boolean
    ' returns true for existing (worth painting at least) buttons
    ' and fills the RECT object we pass in, with the correct dimentions
    ' for itmIdx
    'Debug.Assert itmIdx > 0 And itmIdx <= m_colIcons.Count
    
    m_nAlign = UserControl.Extender.Align
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        Dim nItemH As Long
        nItemH = m_nOptimalHeight - 3
        rItem.Left = UserControl.ScaleLeft + 1
        rItem.Right = UserControl.ScaleWidth - UserControl.ScaleLeft - 1
        If rItem.Right - rItem.Left > 0 Then
            rItem.Top = FIRST_OFFSET + (itmIdx - 1) * (nItemH + STANDARD_OFFSET)
            rItem.Bottom = rItem.Top + nItemH
            If rItem.Bottom > UserControl.ScaleTop + UserControl.ScaleHeight Then
                rItem.Left = 0
                rItem.Right = 0
                rItem.Top = 0
                rItem.Bottom = 0
                ItemRect = False
            Else
                ItemRect = True
            End If
        Else
            rItem.Left = 0
            rItem.Right = 0
            ItemRect = False
        End If
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        rItem.Top = UserControl.ScaleTop + 1
        rItem.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop - 1
        If rItem.Bottom - rItem.Top > 0 Then
            If m_maxCount = 0 Then Exit Function
            Dim nItemW As Long
            nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
            nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
            rItem.Left = UserControl.ScaleLeft + FIRST_OFFSET + (itmIdx - 1) * (nItemW + STANDARD_OFFSET)
            rItem.Right = rItem.Left + nItemW
            ItemRect = True
        Else
            rItem.Top = 0
            rItem.Bottom = 0
            ItemRect = False
        End If
    End If
End Function

'Private Function PointInElement(ByVal x As Single, ByVal y As Single) As Integer
'    'returns index of an element, the point of coordinates
'    'given is in
'    If m_maxCount <> 0 Then
'        m_nAlign = UserControl.Extender.Align
'        Dim nEl As Integer
'        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
'            PointInElement = 0
'            If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
'                Dim nItemH As Long
'                nItemH = m_nOptimalHeight - 3
'
'                nEl = Int((y - UserControl.ScaleTop - FIRST_OFFSET) / (nItemH + STANDARD_OFFSET)) + 1
'                If Not (nEl > m_maxCount Or nEl < 0) Then
'                    If (y - UserControl.ScaleTop - FIRST_OFFSET) - (nEl - 1) * (nItemH + STANDARD_OFFSET) > -2 Then PointInElement = nEl
'                End If
'            End If
'        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
'            PointInElement = 0
'            If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
'
'                Dim nItemW As Long
'
'                nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
'                nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
'
'                nEl = Int((x - UserControl.ScaleLeft - FIRST_OFFSET) / (nItemW + STANDARD_OFFSET)) + 1
'                If Not (nEl > m_maxCount Or nEl < 0) Then
'                    If (x - UserControl.ScaleLeft - FIRST_OFFSET) - (nEl - 1) * (nItemW + STANDARD_OFFSET) > -2 Then PointInElement = nEl
'                End If
'
'            End If
'        End If
'    End If
'
'End Function

Private Function ElementFromPoint(ByVal x As Single, ByVal y As Single) As Integer
    'returns index of an element, the point of coordinates
    'given is in
    
    ' this is my new version of the PointInElement function
    ' renamed to better suite the purpose of the function
    Dim nEl As Integer
    Dim icn As clsIcon
    Dim i As Integer
    
    If m_maxCount <> 0 Then
        ' not 0, we have items.
        m_nAlign = UserControl.Extender.Align
        ' default to 0
        ElementFromPoint = 0
                
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of 0
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        For i = 1 To m_colIcons.Count
            If IsPointInElement(x, y, i) Then
                ElementFromPoint = i
                Exit For
            End If
        Next i
    End If
    
End Function

'Private Function IsPointInElement(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
'    'checks, whether the point is within area of the point
'    'of given index or not
'
'    'returns index of an element, the point of coordinates
'    'given is in
'    m_nAlign = UserControl.Extender.Align
'    Dim nEl As Integer
'    If m_maxCount <> 0 Then
'        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
'            IsPointInElement = False
'            If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
'
'
'                Dim nItemH As Long
'                nItemH = m_nOptimalHeight - 3
'
'                Dim yOffs As Single
'                yOffs = y - UserControl.ScaleTop - FIRST_OFFSET
'
'                IsPointInElement = (y > (idx - 1) * (nItemH + STANDARD_OFFSET)) And (y < idx * (nItemH + STANDARD_OFFSET) - STANDARD_OFFSET)
'            End If
'        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
'            IsPointInElement = False
'            If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
'
'                Dim nItemW As Long
'                nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
'                nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
'
'                Dim xOffs As Single
'                xOffs = x - UserControl.ScaleLeft - FIRST_OFFSET
'
'                IsPointInElement = (x > (idx - 1) * (nItemW + STANDARD_OFFSET)) And (x < idx * (nItemW + STANDARD_OFFSET) - STANDARD_OFFSET)
'            End If
'        End If
'    End If
'End Function

Private Function IsPointInElement(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
    'checks, whether the point is within area of the point
    'of given index or not
    
    ' NEW version of this fuction, easier to read.
    
    Dim nEl As Integer
    Dim oRect As RECT
    
    m_nAlign = UserControl.Extender.Align
    ' default to False
    IsPointInElement = False
    
    If m_maxCount <> 0 Then
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of false
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        If ItemRect(idx, oRect) Then
            IsPointInElement = CBool(PtInRect(oRect, x, y))
        End If
    End If

End Function

Private Sub ActivateWindow(ByVal nEl As Integer)
    ' this function is called when we need to activate
    ' a window because a button on the taskbar was clicked.
    
    ' major re-write of this function also
    ' about half the size it used to be.
    Dim hWnd As Long
    Dim hWndStart As Long
    Dim hWndLast As Long
    Dim sClassName As String
    Dim lLen As Long
    
    If nEl < 1 Or nEl > m_maxCount Then Exit Sub
    On Error GoTo ActivateFailed
    
    Set m_refActive = m_colIcons(nEl)
    If m_refActive.State = vbMinimized Then
        ShowWindow m_refActive.hWnd, SW_SHOW
        ShowWindow m_refActive.hWnd, SW_RESTORE
    Else
        BringWindowToTop m_refActive.hWnd
    End If
    
ActivateFailed:
    ' it may happen
End Sub

Private Sub ClearCollection()
On Error GoTo ClearCollectionError
    ' this makes sure all the windows that we had hidden
    ' are shown again. Just cleaning up on a Usercontrol_Hide.
    m_nIndexBeingSelected = 0
    m_bInsetSelected = False

    Dim icn As clsIcon
    For Each icn In m_colIcons
        If icn.State = vbMinimized Then
            ShowWindow icn.hWnd, SW_SHOW
        End If
    Next
    Exit Sub
ClearCollectionError:
    Debug.Print "Error code: " + Str$(Err.Number) + " in cleanup"
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets the text color for the task bar."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
    UserControl_Paint
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets the BackColor for the task bar (not the buttons)"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    m_BackColor = vNewValue
    UserControl_Paint
    PropertyChanged "BackColor"
End Property

Public Property Get RaisedBackColor() As OLE_COLOR
    RaisedBackColor = m_RaisedBackColor
End Property

Public Property Let RaisedBackColor(ByVal vNewValue As OLE_COLOR)
    m_RaisedBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "RaisedBackColor"
End Property

Public Property Get SunkenBackColor() As OLE_COLOR
Attribute SunkenBackColor.VB_Description = "Sets the BackColor of the button that is currently active or Sunken."
    SunkenBackColor = m_SunkenBackColor
End Property

Public Property Let SunkenBackColor(ByVal vNewValue As OLE_COLOR)
    m_SunkenBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "SunkenBackColor"
End Property

Public Property Get SelectingBackColor() As OLE_COLOR
Attribute SelectingBackColor.VB_Description = "Sets the BackColor of buttons that are currently being selected (clicked/held down with the mouse.)"
    SelectingBackColor = m_SelectingBackColor
End Property
'
Public Property Let SelectingBackColor(ByVal vNewValue As OLE_COLOR)
    m_SelectingBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "SelectingBackColor"
End Property

Public Property Get Style() As enmStyles
Attribute Style.VB_Description = "Sets the style to use when drawing the task bar."
    Style = m_Style
End Property

Public Property Let Style(ByVal vNewValue As enmStyles)
    m_Style = vNewValue
    PropertyChanged "Style"
End Property

Public Property Get CoolBarSeparator() As Boolean
Attribute CoolBarSeparator.VB_Description = "True/False Show a separator bar inbetween the buttons when Style is set to CoolBar."
    CoolBarSeparator = m_CoolBarSeparator
End Property

Public Property Let CoolBarSeparator(ByVal vNewValue As Boolean)
    m_CoolBarSeparator = vNewValue
    PropertyChanged "CoolBarSeparator"
End Property

Public Property Get ShowActive() As Boolean
    ShowActive = m_ShowActive
End Property

Public Property Let ShowActive(ByVal vNewValue As Boolean)
    m_ShowActive = vNewValue
    PropertyChanged "ShowActive"
End Property

Public Property Get AutoHide() As Boolean
Attribute AutoHide.VB_Description = "AutoHide works just like windows taskbar's autohide feature.  When you move the mouse off of the taskbar it hides (after the specified time in AutoHideWait has passed), and when you move back over the smaller version of the bar, it shows again."
    AutoHide = m_AutoHide
End Property

Public Property Let AutoHide(ByVal vNewValue As Boolean)
    m_AutoHide = vNewValue
    PropertyChanged "AutoHide"
End Property

Public Property Get AutoHideWait() As Integer
Attribute AutoHideWait.VB_Description = "This is the # of milliseconds (1000 = 1 second) to wait until hiding the bar, after moving off of it when AutoHide = True."
    AutoHideWait = m_AutoHideWait
End Property

Public Property Let AutoHideWait(ByVal vNewValue As Integer)
    m_AutoHideWait = vNewValue
    PropertyChanged "AutoHideWait"
End Property

Public Property Get AutoHideAnimate() As Boolean
Attribute AutoHideAnimate.VB_Description = "If AutoHide is true, and you would like the bar to Slide out of side, instead of just disappearing, then set this to true."
    AutoHideAnimate = m_AutoHideAnimate
End Property

Public Property Let AutoHideAnimate(ByVal vNewValue As Boolean)
    m_AutoHideAnimate = vNewValue
    PropertyChanged "AutoHideAnimate"
End Property

Public Property Get AutoHideAnimateFrames() As Integer
Attribute AutoHideAnimateFrames.VB_Description = "This sets how many frames to use when animating the AutoHide feature."
    AutoHideAnimateFrames = m_AutoHideAnimateFrames
End Property

Public Property Let AutoHideAnimateFrames(ByVal vNewValue As Integer)
    m_AutoHideAnimateFrames = vNewValue
    PropertyChanged "AutoHideAnimateFrames"
End Property
