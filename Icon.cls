VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' an object of this class represents a window
' icon displayed in application task bar

' the concept of 'touch' is for iterating through
' a collection of objects of the clsIcon type

' unyouched items will be then marked (for delete ?)


Private m_strTitle As String ' holds the caption from the window
Private m_hWnd As Long ' the windows handle
Private m_nState As FormWindowStateConstants  ' minimized/maximozed/normal
Private m_hIcon As Long ' handle to the icon
' m_ref_Icons refers to icon to be presented

Private m_bTouch As Boolean
Private m_bNew As Boolean ' is the icon new?
Private m_bChanged As Boolean ' has the window changed?

Public Property Get IsNew() As Boolean
    IsNew = m_bNew
End Property

Public Property Get Title() As String
    Title = m_strTitle
End Property

Public Property Let Title(ByVal strNewValue As String)
    m_strTitle = strNewValue
    m_bChanged = True
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal hNewValue As Long)
    m_hWnd = hNewValue
End Property

Public Property Get State() As FormWindowStateConstants
    State = m_nState
End Property

Public Property Let State(ByVal newStateValue As FormWindowStateConstants)
    m_nState = newStateValue
End Property

Friend Sub ClearTouch()
    m_bNew = False
    m_bTouch = False
    m_bChanged = False
End Sub

Friend Sub Touch()
    m_bTouch = True
End Sub

Friend Function IsTaught() As Boolean
    IsTaught = m_bTouch
End Function

Private Sub Class_Initialize()
    m_bNew = True
    m_bTouch = True
    m_bChanged = True
End Sub

Public Property Get IsChanged() As Boolean
    IsChanged = m_bChanged
End Property

Public Property Let IconPtr(ByVal hIcon As Long)
    m_hIcon = hIcon
End Property

Public Property Get IconPtr() As Long
    'IconPtr = m_hIcon
    
    ' this gets the icon's handle and creates a local copy
    ' so we can use it
    Dim hIcon As Long
    Dim hLocalIcon As Long
    If hWnd <> 0 Then
        
        hIcon = GetWndIcon(m_hWnd, False)
        hLocalIcon = CopyIcon(hIcon)
        IconPtr = hLocalIcon
        
        If IconPtr <> 0 And m_hIcon <> 0 Then
            ' we had an icon already and are now replacing it.
            ' so we have to destroy the old one
            DestroyIcon m_hIcon
        End If
        
        ' now set the new one
        m_hIcon = hLocalIcon
    Else
        IconPtr = 0
        m_hIcon = 0
    End If
End Property

Private Sub Class_Terminate()
    ' destroy our local copy of the icon ...
    If m_hIcon <> 0 Then
        DestroyIcon m_hIcon
    End If
End Sub
