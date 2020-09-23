Attribute VB_Name = "Declares"
Option Explicit

Global ghWnd As Long

Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CYCAPTION = 4
'
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_MONO = &H8000     ' For monochrome borders.
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_CENTER = &H1
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9

Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDITILE = &H226
Public Const WM_SIZE = &H5
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETFOCUS = &H7

Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDISETMENU = &H230

Public Const WM_CLOSE = &H10

Const GWL_WNDPROC = (-4)
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function OleTranslateColor Lib "Olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hpal As Long, pcolorref As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Const CLR_INVALID = &HFFFF

Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private objTB As TaskBar
Private pfWndProc As Long
    
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SendMessageTimeout Lib "user32" Alias _
        "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As _
        Long, ByVal wParam As Long, ByVal lParam As Long, ByVal _
        fuFlags As Long, ByVal uTimeout As Long, lpdwResult As _
        Long) As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GCL_HICON = (-14)
Public Const GCL_HICONSM = (-34)
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Public Const DI_IMAGE = &H2
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8
Public Const DI_MASK = &H1
Public Const DI_NORMAL = &H3
Public Const WM_ACTIVATE = &H6
Public Const SIZE_MINIMIZED = 1


Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal _
    uFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal _
    hWnd As Long, ByVal prcRect As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_CLOSE = &HF060
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function ShowSystemMenu(ByVal hWnd As Long)

    Dim curpos As POINTAPI  ' holds the current mouse coordinates
    Dim retval As Long        ' generic return value
    Dim lMenu As Long
    Dim lSys As Long
    
    ' get a copy of the system menu from the window.
    lSys = GetSystemMenu(hWnd, 0)
    
    retval = GetCursorPos(curpos)
    ' raise the menu at the current cursor position
    lMenu = TrackPopupMenu(lSys, TPM_RETURNCMD Or TPM_LEFTBUTTON Or TPM_TOPALIGN Or TPM_LEFTALIGN, curpos.x, curpos.y, 0, hWnd, 0)
    
    ' handle menu clicks
    Select Case lMenu
        Case 61456
            ' move
            DefWindowProc hWnd, WM_SYSCOMMAND, SC_MOVE, 0
        Case 61536 ' close
            PostMessage hWnd, WM_CLOSE, 0&, 0&
        Case 61504 ' next mdi child
            SendMessage ghWnd, WM_MDINEXT, hWnd, 0&
        Case 61440
            ' Size
            DefWindowProc hWnd, WM_SYSCOMMAND, SC_SIZE, 0
        Case 61472
            ' Minimize
            ShowWindow hWnd, 2
        Case 61488
            ' Maximize
            SendMessage ghWnd, WM_MDIMAXIMIZE, hWnd, 0&
        Case 61728
            ' restore
            SendMessage ghWnd, WM_MDIRESTORE, hWnd, 0&
        Case Else
            'Debug.Print "Menu: " & lMenu & vbCrLf
    End Select
End Function

Public Function GetWndIcon(hWndIcon As Long, bLarge As Boolean) As Long
    'Attempts to grab the icon for a window
    
    ' First off, attempt WM_GETICON, use SendMesageTimeout so we don't
    '  hang on windows that aren't responding
    SendMessageTimeout hWndIcon, WM_GETICON, IIf(bLarge, ICON_BIG, _
                       ICON_SMALL), 0, 0, 1000, GetWndIcon
                
    If GetWndIcon = 0 Then
        ' If WM_GETICON didn't return anything, try using
        '  GetClassLong to get the icon for the window's class
        GetWndIcon = GetClassLong(hWndIcon, IIf(bLarge, GCL_HICON, _
                   GCL_HICONSM))
    End If
    
    If GetWndIcon = 0 Then
        
        'GetWndIcon = LoadIcon(0&, )
    End If
End Function


Public Sub SubClassParentWnd(ByRef obj As TaskBar)
    ' purpose of the function is to substitutue
    ' WndProc of MDIClient window
    
    Dim hWnd As Long
    ghWnd = FindWindowEx(GetParent(obj.hWnd), 0, "MDIClient", vbNullString)
    If GetParent(obj.hWnd) <> 0 Then
        hWnd = GetParent(obj.hWnd)
        Set objTB = obj
        pfWndProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf MDI_ParentWndProc)
    End If
End Sub

Public Sub UnSubClassParentWnd(ByRef obj As TaskBar)
    ' to revert the previous state
    
    If GetParent(obj.hWnd) <> 0 Then
        SetWindowLong ghWnd, GWL_WNDPROC, pfWndProc
        Set objTB = Nothing
    End If
End Sub

Function MDI_ParentWndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lret As Long
    Dim lWinStyle As Long
    ' since message is handled, we can notify
    ' an object something interesting happened
    lret = CallWindowProc(pfWndProc, hWnd, msg, wParam, lParam)
        
    If GetParent(objTB.hWnd) <> 0 Then
        If msg = WM_MDIGETACTIVE Then
            objTB.OnRefresh lret
            objTB.RaiseChildActivate lret
        ElseIf msg = WM_KILLFOCUS Then
            objTB.OnRefresh wParam
        ElseIf msg = WM_MDIRESTORE Then
            objTB.RaiseChildRestore wParam
        ElseIf msg = WM_MDIMAXIMIZE Then
            objTB.RaiseChildMaximize wParam
        ElseIf msg = WM_MDICREATE Then
            objTB.RaiseChildCreate lret
        ElseIf msg = WM_MDIDESTROY Then
            objTB.OnRefresh
            objTB.RaiseChildDestroy wParam
        'Else

            'Debug.Print "hWnd: " & hWnd & " :msg: " & msg & " :wParam: " & wParam & " :lParam: " & lParam & " :lret: " & lret & " :WindowText(hWnd): " & WindowText(hWnd) & " :WindowText(lret): " & WindowText(lret) & " :WindowText(wParam): " & WindowText(wparam)
        End If


    End If

    MDI_ParentWndProc = lret
End Function

Public Function GetElement(ByVal hWnd As Long) As clsIcon
    Dim i As Integer
    For i = 0 To objTB.m_colIcons.Count
        If objTB.m_colIcons(i).hWnd = hWnd Then
            Set GetElement = objTB.m_colIcons(i)
            Exit For
        End If
    Next i
End Function

Public Function TranslateColor(ByVal clrColor As OLE_COLOR, _
    Optional hPalette As Long = 0) As Long

    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
      TranslateColor = CLR_INVALID
    End If

End Function

Public Function WindowText(ByVal hWnd As Long) As String
    ' this function returns the caption for the window
    ' specified by hWnd
    Dim sCaption As String
    Dim nRet As Long
    
    sCaption = Space$(256)
    nRet = GetWindowText(hWnd, sCaption, Len(sCaption))
    If nRet Then
        sCaption = Left(sCaption, nRet)
    End If
    
    WindowText = sCaption
End Function
