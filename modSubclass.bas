Attribute VB_Name = "modSubclass"
Option Explicit

'API declares / constants / types -------------------------------------------
Public Const GWL_WNDPROC = -4

Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                        (lpvDest As Any, _
                                        lpvSource As Any, _
                                        ByVal cbCopy As Long)

Public Declare Function GetProp Lib "user32" Alias "GetPropA" _
                                        (ByVal hwnd As Long, _
                                        ByVal lpString As String) _
                                        As Long

Public Declare Function SetProp Lib "user32" Alias "SetPropA" _
                                        (ByVal hwnd As Long, _
                                        ByVal lpString As String, _
                                        ByVal hData As Long) _
                                        As Long

Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
                                        (ByVal hwnd As Long, _
                                        ByVal lpString As String) _
                                        As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                        (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) _
                                        As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                                        (ByVal lpPrevWndFunc As Long, _
                                        ByVal hwnd As Long, _
                                        ByVal Msg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) _
                                        As Long
'End of API -----------------------------------------------------------------

Public Const WNDPROPERTY As String = "AutoScroll" 'Property name to be
                                                  'attached to the
                                                  'subclassed window

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) _
                           As Long
'This procedure replaces the original Windows procedure.

    'Get back to the class to process the Window message:
    WindowProc = ObjectFromPointer(hwnd).WindowProc(hwnd, uMsg, _
                                                    wParam, lParam)

End Function

Private Function ObjectFromPointer(ByVal hwnd As Long) As AutoScroll
'Obtains the AutoScroll object from its pointer.

    Dim lngObjectPointer As Long, objAutoScroll As AutoScroll
    
    'Get the pointer for the AutoScroll object stored in the subclassed
    'window as a property:
    lngObjectPointer = GetProp(hwnd, WNDPROPERTY)
    
    If lngObjectPointer Then
        'Copy the contents at the pointer into the temp. AutoScroll object:
        CopyMemory objAutoScroll, lngObjectPointer, 4&
        
        'Return the object (i.e. Assign to legal reference):
        Set ObjectFromPointer = objAutoScroll
        
        'Destroy the temp. object (illegal reference):
        CopyMemory objAutoScroll, 0&, 4&
    End If

End Function
