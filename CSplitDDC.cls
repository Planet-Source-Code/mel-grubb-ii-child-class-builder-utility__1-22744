VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSplitDDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' SplitDDC - A splitter class using the Desktop window to draw a splitter bar,
'   therefore allowing splitting of MDI forms as well as standard forms.
'
' Version   Date        User            Notes
'   1.0     7/7/98      Steve McMahon   Initial version
'   1.1     11/16/00    Mel Grubb II
'               * Change mouse events to follow standard of allowing the splitter
'                 object itself to redirect mouse events instead of the parent form.
'               * Eliminated parent form mouse clipping region
'               * Added support for custom mouse pointers
'               * Rearranged and standardized code formatting
'   1.2     11/29/00    Mel Grubb II    Added error handlers
'                                       Applied new coding standards
'===============================================================================
Option Explicit

'===============================================================================
' Types
'===============================================================================
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'===============================================================================
' API Declarations and related constants
'===============================================================================
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
    Private Const R2_BLACK = 1
    Private Const R2_COPYPEN = 13
    Private Const R2_LAST = 16
    Private Const R2_MASKNOTPEN = 3
    Private Const R2_MASKPEN = 9
    Private Const R2_MASKPENNOT = 5
    Private Const R2_MERGENOTPEN = 12
    Private Const R2_MERGEPEN = 15
    Private Const R2_MERGEPENNOT = 14
    Private Const R2_NOP = 11
    Private Const R2_NOT = 6
    Private Const R2_NOTCOPYPEN = 4
    Private Const R2_NOTMASKPEN = 8
    Private Const R2_NOTMERGEPEN = 2
    Private Const R2_NOTXORPEN = 10
    Private Const R2_WHITE = 16
    Private Const R2_XORPEN = 7
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Const SM_CXBORDER = 5
    Private Const SM_CYBORDER = 6
    Private Const SM_CYCAPTION = 4
    Private Const SM_CYMENU = 15
Private Declare Function GetActiveWindow Lib "user32" () As Long

'===============================================================================
' Constants
'===============================================================================
Private Const mc_strModuleID = "SplitDDC."

'===============================================================================
' Member Variables
'===============================================================================
Private m_blnDrawing As Boolean                 ' Whether or not we are in draw mode
Private m_rcCurrent As RECT                     ' The ractangle we are currently drawing
Private m_rcWindow As RECT                      ' The rectangle defined by the parent window
Private m_rcSplitterObject As RECT              ' The rectangle defined by the splitter object

Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants
Private m_lngBorder(1 To 4) As Long
Private m_objSplitter As Object                 ' The splitter object represents the splitter visually, usually a picturebox
Private m_bIsMDI As Boolean
Private m_bSplitting As Boolean

'===============================================================================
' Enumerations
'===============================================================================
Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum

Public Enum ESplitBorderTypes
   espbLeft = 1
   espbTop = 2
   espbRight = 3
   espbBottom = 4
End Enum


'===============================================================================
' SplitObject - The object to be used as the splitter bar.  This is usually a
'   picturebox, but it could be anything with an hWnd
'===============================================================================
Public Property Get SplitObject() As Object
    On Error GoTo ErrorHandler

    Set SplitObject = m_objSplitter
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "SplitObject_Get"

End Property
Public Property Let SplitObject(ByRef objSplitter As Object)
    On Error Resume Next

    Set m_objSplitter = objSplitter
    With objSplitter
        .BorderStyle = 0
        If (m_eOrientation = espHorizontal) Then
            ' Horizontal splitter
            If .MousePointer <> 99 Then .MousePointer = vbSizeNS
        Else
            ' Vertical splitter
            If .MousePointer <> 99 Then .MousePointer = vbSizeWE
        End If
    End With
End Property


'===============================================================================
' Border - Defines a buffer around the area in which the splitter is allowed to
'   operate
'===============================================================================
Public Property Let Border(ByVal eBorderType As ESplitBorderTypes, ByVal lngSize As Long)
    On Error GoTo ErrorHandler
    
    m_lngBorder(eBorderType) = lngSize
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Border_Let"

End Property
Public Property Get Border(ByVal eBorderType As ESplitBorderTypes) As Long
    On Error GoTo ErrorHandler
    
    Border = m_lngBorder(eBorderType)
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Border_Get"

End Property


'===============================================================================
' Orientation - Whether this splitter is horizontal or vertical
'===============================================================================
Public Property Get Orientation() As eOrientationConstants
    On Error GoTo ErrorHandler
    
    Orientation = m_eOrientation
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Orientation_Get"

End Property
Public Property Let Orientation(ByVal EOrientation As eOrientationConstants)
    On Error GoTo ErrorHandler
    
    m_eOrientation = EOrientation
    If Not (m_objSplitter Is Nothing) Then
        With m_objSplitter
            If (m_eOrientation = espHorizontal) Then
                If .MousePointer <> 99 Then .MousePointer = vbSizeNS
                m_lngBorder(espbTop) = 64
                m_lngBorder(espbBottom) = 64
                m_lngBorder(espbLeft) = 0
                m_lngBorder(espbRight) = 0
            Else
                If .MousePointer <> 99 Then .MousePointer = vbSizeWE
                m_lngBorder(espbTop) = 0
                m_lngBorder(espbBottom) = 0
                m_lngBorder(espbLeft) = 64
                m_lngBorder(espbRight) = 64
            End If
        End With
    End If
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Orientation"

End Property


'===============================================================================
' SplitterMouseDown - Send the splitter object's mousedown events here
'===============================================================================
Public Sub SplitterMouseDown(ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    Dim tP As POINTAPI

    ' Set the reference to the owner form
    m_hWnd = hwnd

    ' Send subsequent mouse messages to the splitter object
    SetCapture m_objSplitter.hwnd

    ' Get the window rectangle on the desktop of the owner window:
    GetWindowRect m_hWnd, m_rcWindow
    ' Get the window rectangle of the splitter object, we'll use this to figure bounds
    GetWindowRect m_objSplitter.hwnd, m_rcSplitterObject

    ' Check if this is an MDI form:
    If (ClassName(m_hWnd) = "ThunderMDIForm") Then
        ' Get the inside portion of the MDI form:
        ' I'm assuming you have a caption,menu and border in your MDI here
        m_rcWindow.Left = m_rcWindow.Left + GetSystemMetrics(SM_CXBORDER)
        m_rcWindow.Right = m_rcWindow.Right - GetSystemMetrics(SM_CXBORDER)
        m_rcWindow.Bottom = m_rcWindow.Bottom - GetSystemMetrics(SM_CYBORDER)
        m_rcWindow.Top = m_rcWindow.Top + GetSystemMetrics(SM_CYBORDER) * 3 _
                                    + GetSystemMetrics(SM_CYCAPTION) _
                                    + GetSystemMetrics(SM_CYMENU)
        m_bIsMDI = True
    Else
        ' Get the client rectangle of the window in screen coordinates:
        GetClientRect m_hWnd, m_rcWindow
        tP.X = m_rcWindow.Left
        tP.Y = m_rcWindow.Top
        ClientToScreen m_hWnd, tP
        m_rcWindow.Left = tP.X
        m_rcWindow.Top = tP.Y
        tP.X = m_rcWindow.Right
        tP.Y = m_rcWindow.Bottom
        ClientToScreen m_hWnd, tP
        m_rcWindow.Right = tP.X
        m_rcWindow.Bottom = tP.Y
        m_bIsMDI = False
    End If
    m_blnDrawing = True  '// start actual drawing from next move message
    
    m_rcCurrent.Left = 0
    m_rcCurrent.Top = 0
    m_rcCurrent.Right = 0
    m_rcCurrent.Bottom = 0
    
    X = (m_objSplitter.Left + X) \ Screen.TwipsPerPixelX
    Y = (m_objSplitter.Top + Y) \ Screen.TwipsPerPixelY
    SplitterMouseMove X, Y
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "SplitterMouseDown"

End Sub


'===============================================================================
' SplitterMouseMove - Send the form containing the splitter's mouse moves
'   here
'===============================================================================
Public Sub SplitterMouseMove(ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    Dim hdc As Long
    Dim rcNew As RECT
    Dim tP As POINTAPI
    
    If (m_blnDrawing) Then
        ' Draw two rectangles in the screen DC to cause splitting:
        ' First get the Desktop DC:
        hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Set the draw mode to XOR:
        SetROP2 hdc, R2_NOT 'XORPEN
        ' Draw over and erase the old rectangle
        ' (if this is the first time, all the coords will be 0 and nothing will get drawn):
        Rectangle hdc, m_rcCurrent.Left, m_rcCurrent.Top, m_rcCurrent.Right, m_rcCurrent.Bottom

        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos tP

        ' Determine where to draw the splitter:
        If (m_eOrientation = espHorizontal) Then
            ' Use m_rcSplitterObject's left and right, and m_rcWindow's top and bottom
            rcNew.Left = m_rcSplitterObject.Left
            rcNew.Right = m_rcSplitterObject.Right
            If (tP.Y >= m_rcWindow.Top + m_lngBorder(espbTop)) And (tP.Y < m_rcWindow.Bottom - m_lngBorder(espbBottom)) Then
                rcNew.Top = tP.Y - 2
                rcNew.Bottom = tP.Y + 2
            Else
                If (tP.Y < m_rcWindow.Top + m_lngBorder(espbTop)) Then
                    rcNew.Top = m_rcWindow.Top + m_lngBorder(espbTop) - 2
                    rcNew.Bottom = rcNew.Top + 5
                Else
                    rcNew.Top = m_rcWindow.Bottom - m_lngBorder(espbBottom) - 2
                    rcNew.Bottom = rcNew.Top + 5
                End If
            End If
        Else
            ' Use m_rcSplitterObject's top and bottom, and m_rcWindow's left and right
            rcNew.Top = m_rcSplitterObject.Top
            rcNew.Bottom = m_rcSplitterObject.Bottom
            If (tP.X >= m_rcWindow.Left + m_lngBorder(espbLeft)) And (tP.X <= m_rcWindow.Right - m_lngBorder(espbRight)) Then
                ' The mouse is within the allowable bounds, define the new rectangle as 4 pixels, centered on the mouse position
                rcNew.Left = tP.X - 2
                rcNew.Right = tP.X + 2
            Else
                ' The mouse is out of bounds, define the new rectangle at the edge of the bounds
                If (tP.X < m_rcWindow.Left + m_lngBorder(espbLeft)) Then
                    rcNew.Left = m_rcWindow.Left + m_lngBorder(espbLeft) - 2
                    rcNew.Right = rcNew.Left + 5
                Else
                    rcNew.Left = m_rcWindow.Right - m_lngBorder(espbRight) - 2
                    rcNew.Right = rcNew.Left + 5
                End If
            End If
        End If
        
        ' Draw the new rectangle
        Rectangle hdc, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        
        ' Store this position so we can erase it next time:
        LSet m_rcCurrent = rcNew
        
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        DeleteDC hdc
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "SplitterMouseDown"

End Sub


'===============================================================================
' SplitterMouseUp -
'===============================================================================
Public Function SplitterMouseUp(ByVal X As Long, ByVal Y As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim hdc As Long
    Dim tP As POINTAPI
    
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (m_blnDrawing) Then
        m_blnDrawing = False
        
        ' Release mouse capture:
        ReleaseCapture
        
        ' Get the Desktop DC:
        hdc = CreateDCAsNull("DISPLAY", 0, 0, 0)
        ' Set to XOR drawing mode:
        SetROP2 hdc, R2_NOT 'XORPEN
        ' Erase the last rectangle:
        Rectangle hdc, m_rcCurrent.Left, m_rcCurrent.Top, m_rcCurrent.Right, m_rcCurrent.Bottom
        ' Clear up the desktop DC:
        DeleteDC hdc
        
        ' Here we ensure the splitter is within bounds before releasing:
        GetCursorPos tP
        
        If (tP.X < m_rcWindow.Left + m_lngBorder(espbLeft)) Then
            tP.X = m_rcWindow.Left + m_lngBorder(espbLeft)
        End If
        If (tP.X > m_rcWindow.Right - m_lngBorder(espbRight)) Then
            tP.X = m_rcWindow.Right - m_lngBorder(espbRight)
        End If
        If (tP.Y < m_rcWindow.Top + m_lngBorder(espbTop)) Then
            tP.Y = m_rcWindow.Top + m_lngBorder(espbTop)
        End If
        If (tP.Y > m_rcWindow.Bottom - m_lngBorder(espbBottom)) Then
            tP.Y = m_rcWindow.Bottom - m_lngBorder(espbBottom)
        End If
        ScreenToClient m_hWnd, tP
        
        ' Move the splitter to the validated final position:
        If (m_eOrientation = espHorizontal) Then
            ' Move the splitter vertically
            m_objSplitter.Top = (tP.Y - 2) * Screen.TwipsPerPixelY
        Else
            ' Move the splitter horizontally
            m_objSplitter.Left = (tP.X - 2) * Screen.TwipsPerPixelX
        End If
        
        ' Return true to tell the owner we have completed splitting:
        SplitterMouseUp = True
    End If
    Exit Function

ErrorHandler:
    ProcessError Err, mc_strModuleID & "SplitterMouseUp"

End Function


'===============================================================================
' Class_Initialize - Set up some defaults
'===============================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    m_eOrientation = espVertical
    m_lngBorder(espbLeft) = 64
    m_lngBorder(espbRight) = 64
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Class_Initialize"

End Sub


'===============================================================================
' ClassName - Return the class name of a given window
'===============================================================================
Private Function ClassName(ByVal lHwnd As Long) As String
    On Error GoTo ErrorHandler
    Dim lLen As Long
    Dim sBuf As String
    
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
    Exit Function

ErrorHandler:
    ProcessError Err, mc_strModuleID & "ClassName"

End Function

