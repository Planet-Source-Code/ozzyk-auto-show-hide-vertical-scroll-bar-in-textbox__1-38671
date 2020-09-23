VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API declares / types / constants -------------------------------------------

'TextBox message constants
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETRECT = &HB2
Private Const WM_NCPAINT = &H85
Private Const WM_GETFONT = &H31

'TextBox windows messages to be intercepted
Private Const WM_KEYUP = &H101
Private Const WM_SIZE = &H5

'Scroll bar orientation constants
Private Const SB_HORZ = 0
Private Const SB_VERT = 1

'Rectangle structure
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Size structure
Private Type SIZE
    cx As Long
    cy As Long
End Type

'API Functions
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, _
                                        ByVal wBar As Long, _
                                        ByVal bShow As Long) _
                                        As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                        ByVal wMsg As Long, _
                                        ByVal wParam As Long, _
                                        lParam As Any) _
                                        As Long

Private Declare Function SelectObject Lib "gdi32" _
                                        (ByVal hDC As Long, _
                                        ByVal hObject As Long) _
                                        As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" _
                                        Alias "GetTextExtentPoint32A" _
                                        (ByVal hDC As Long, _
                                        ByVal lpsz As String, _
                                        ByVal cbString As Long, _
                                        lpSize As SIZE) _
                                        As Long

Private Declare Function ReleaseDC Lib "user32" _
                                        (ByVal hwnd As Long, _
                                        ByVal hDC As Long) _
                                        As Long

Private Declare Function GetDC Lib "user32" _
                                        (ByVal hwnd As Long) _
                                        As Long
'End of API -----------------------------------------------------------------

'Member variables:
Private mlngWindow As Long 'Handle to the target window
Private mlngProcOld As Long 'Pointer to the default Windows procedure
Private mblnSBVisible As Boolean 'State of the scroll bar

Public Property Get Window() As Long

    Window = mlngWindow

End Property

Public Property Let Window(ByVal hwnd As Long)

    If IsWindow(hwnd) Then
        If mlngWindow Then 'If there's another subclassed window
            Call Cleanup
        End If

        mlngWindow = hwnd
        
        'Subclass this window:
        SubclassWindow mlngWindow
    Else
        Err.Raise Number:=vbObjectError + 512, _
                  Description:="Invalid window handle."
        
        mlngWindow = 0
    End If

End Property

Friend Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) _
                           As Long

    'Let the original windows procedure handle the message first:
    WindowProc = CallWindowProc(mlngProcOld, hwnd, uMsg, wParam, lParam)

    Select Case uMsg
        Case WM_KEYUP, WM_SIZE
            'Compare the height of the text to the height of the
            'client area:
            mblnSBVisible = (GetTextHeight > GetClientHeight)
            
            ShowOrHideScrollBar mblnSBVisible
    End Select

End Function

Private Sub SubclassWindow(ByVal hwnd As Long)

    If GetProp(hwnd, WNDPROPERTY) Then Exit Sub
    
    If SetProp(hwnd, WNDPROPERTY, ObjPtr(Me)) Then 'If the object pointer
                                                   'can be attached to the
                                                   'window as a property
        mlngProcOld = SetWindowLong(hwnd, GWL_WNDPROC, _
                                    AddressOf modSubclass.WindowProc)
    End If

End Sub

Private Sub Cleanup()

    If IsWindow(mlngWindow) Then
        'Un-subclass:
        If mlngProcOld Then
            'Restore the original Windows procedure:
            SetWindowLong mlngWindow, GWL_WNDPROC, mlngProcOld
            
            'Remove the added property:
            RemoveProp mlngWindow, WNDPROPERTY

            'Reset variables:
            mlngProcOld = 0
            mlngWindow = 0
        End If
    End If

End Sub

Private Sub ShowOrHideScrollBar(ByVal State As Boolean)

    ShowScrollBar mlngWindow, SB_VERT, State

    If State Then
        'If on XP, or a skin/theme is otherwise applied to the window _
        'scroll bar may not be painted correctly. So, force the
        'non-client area to repaint:
        SendMessage mlngWindow, WM_NCPAINT, 0&, ByVal 0&
    End If
    
End Sub

Private Function GetClientHeight() As Long

    Dim udtRect As RECT
    
    SendMessage mlngWindow, EM_GETRECT, 0&, udtRect
    
    With udtRect
        GetClientHeight = .Bottom - .Top
    End With

End Function

Private Function GetTextHeight() As Long

    Dim lngFontHandle As Long
    Dim lngOldFontHandle As Long
    Dim lngDeviceContext As Long
    Dim strTemp As String
    Dim udtSize As SIZE
    Dim lngCharHeight As Long
    Dim lngLineCount As Long
    
    'Find the height of characters in TextBox's font:
    lngDeviceContext = GetDC(mlngWindow)
    
    lngFontHandle = SendMessage(mlngWindow, WM_GETFONT, 0&, ByVal 0&)
    
    lngOldFontHandle = SelectObject(lngDeviceContext, lngFontHandle)
    
    strTemp = "ABC"

    GetTextExtentPoint32 lngDeviceContext, strTemp, Len(strTemp), udtSize
    
    lngCharHeight = udtSize.cy
    
    'Clean up:
    Call SelectObject(lngDeviceContext, lngOldFontHandle)
    Call ReleaseDC(mlngWindow, lngDeviceContext)
    
    'Find the height of the text by multiplying character height by the
    'line count:
    lngLineCount = SendMessage(mlngWindow, EM_GETLINECOUNT, 0&, ByVal 0&)
    
    GetTextHeight = lngCharHeight * lngLineCount

End Function

Private Sub Class_Terminate()

    Call Cleanup

End Sub
