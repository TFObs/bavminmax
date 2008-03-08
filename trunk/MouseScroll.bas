Attribute VB_Name = "MouseScroll"
Option Explicit
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Public Declare Function SetWindowsHookEx Lib "user32" Alias _
    "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
    ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
    ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long


Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type MOUSEHOOKSTRUCT
  pt As POINTAPI
  hWnd As Long
  wHitTestCode As Long
  dwExtraInfo As Long
End Type

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

Public Const WH_MOUSE = 7
Private Const WHEEL_DELTA = 120

Public Const GWL_WNDPROC = -4

Public hook As Long
Dim nKeys As Long, Delta As Long, XPos As Long, YPos As Long
Dim OriginalWindowProc As Long

Public Enum mButtons
  LBUTTON = &H1
  MBUTTON = &H10
  RBUTTON = &H2
End Enum

Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, _
                          lParam As MOUSEHOOKSTRUCT) As Long
    Select Case nCode
      Case Is < 0
        MouseProc = CallNextHookEx(hook, nCode, wParam, lParam)
      Case 0
        If lParam.hWnd = frmHaupt.hWnd Then
          Select Case wParam
            Case WM_MBUTTONDOWN
              MouseWheelDown lParam.pt.x, lParam.pt.y
              Debug.Print "Button down:" & lParam.pt.x & "," & lParam.pt.y
            Case WM_MBUTTONUP
              MouseWheelUp lParam.pt.x, lParam.pt.y
              Debug.Print "Button up:" & lParam.pt.x & "," & lParam.pt.y
          End Select
        End If
    End Select
End Function

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) _
                           As Long
    Select Case uMsg
      Case WM_MOUSEWHEEL
        nKeys = wParam And 65535
        Delta = wParam / 65536 / WHEEL_DELTA
        XPos = lParam And 65535
        YPos = lParam / 65536

        MouseWheelRotation Delta, nKeys, XPos, YPos, hWnd
        Debug.Print "Mousewheel at (" & XPos & "," & YPos & ") Delta:" & _
                    Delta & "  Keys:" & nKeys
    End Select

    WindowProc = CallWindowProc(OriginalWindowProc, hWnd, uMsg, wParam, _
                                lParam)
End Function

'Nicht vergessen: Ende() ausführen!!!
Public Function MInit(Form As Form)
    hook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, _
                            GetCurrentThreadId)
    OriginalWindowProc = SetWindowLong(Form.hWnd, GWL_WNDPROC, _
                                       AddressOf WindowProc)
End Function

Public Function Mende()
    UnhookWindowsHookEx hook
    SetWindowLong frmHaupt.hWnd, GWL_WNDPROC, OriginalWindowProc
End Function

Public Function MouseWheelRotation(Richtung As Long, Buttons As mButtons, _
                                   x As Long, y As Long, hWnd As Long)
    'Hier die eigene Auswertung rein
    'If TypeOf frmHaupt.ActiveControl Is MSHFlexGrid Then
      If Richtung = 1 Then
        frmHaupt.ScrollUp
      ElseIf Richtung = -1 Then
        frmHaupt.ScrollDown
      End If
    'End If
End Function

Public Function MouseWheelUp(x As Long, y As Long)
'frmHaupt.Label5.Caption = "WheelButtonUp"
End Function

Public Function MouseWheelDown(x As Long, y As Long)
'frmHaupt.Label5.Caption = "WheelButtonDown"
End Function





