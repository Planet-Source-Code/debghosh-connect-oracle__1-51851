Attribute VB_Name = "txtSubclass"
'Subclass A TextBox
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = -4
    Private Const WM_RBUTTONDOWN As Long = &H204
    Public Const WM_RBUTTONUP = &H205
    Private Const WM_COPY As Long = &H301
    Private Const WM_PASTE As Long = &H302
    Global lpPrevWndProc As Long
    Global txt As TextBox
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Sub Hook()
    lpPrevWndProc = SetWindowLong(txt.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHook()
    Dim lngRValue As Long
    lngRValue = SetWindowLong(txt.hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Select Case uMsg
            Case WM_RBUTTONDOWN
                MsgBox "Subclassing TextBox. Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", vbCritical
            Case WM_COPY
                MsgBox "Subclassing TextBox. Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", vbCritical
            Case WM_PASTE
                MsgBox "Subclassing TextBox. Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", vbCritical
        End Select
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
