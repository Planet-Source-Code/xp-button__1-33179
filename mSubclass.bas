Attribute VB_Name = "mSubclass"
Option Explicit

Private Const GWL_WNDPROC = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Sub HookForm(ByVal hWnd As Long)
   Dim procOld As Long
   Dim r As Long
   
   procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf FormProc)
   If procOld Then
      r = SetProp(hWnd, hWnd, procOld)
   End If
End Sub

Public Sub HookControl(ByVal hWnd As Long)
   Dim procOld As Long
   Dim r As Long
   
   procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ControlProc)
   If procOld Then
      r = SetProp(hWnd, hWnd, procOld)
   End If
End Sub

Public Sub UnHook(ByVal hWnd As Long)
   Dim procOld As Long
   Dim r As Long

   procOld = GetProp(hWnd, hWnd)
   If procOld Then
      r = SetWindowLong(hWnd, GWL_WNDPROC, procOld)
      r = RemoveProp(hWnd, hWnd)
   End If
End Sub

Private Function FormProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim procOld As Long
   Dim DIS As DRAWITEMSTRUCT
   Dim State As Long
   Dim bPushed As Boolean
   Dim bFocus As Boolean
   Dim bEnabled As Boolean
   Dim bHover As Boolean
   Dim Ctrl As Control
   
   If uMsg = WM_DRAWITEM Then
      CopyMemory DIS, ByVal lParam, Len(DIS)
      If DIS.CtlType = ODT_BUTTON Then
         Set Ctrl = Form1.GetControl(DIS.hWndItem)
         If Not Ctrl Is Nothing Then
            State = SendMessage(DIS.hWndItem, BM_GETSTATE, 0&, 0&)
            bPushed = (State And BST_PUSHED)
            bFocus = (State And BST_FOCUS)
            bEnabled = IsWindowEnabled(DIS.hWndItem)
            State = SendMessage(DIS.hWndItem, BM_GETCHECK, 0&, 0&)
            bHover = GetProp(DIS.hWndItem, "Hover")
            If TypeOf Ctrl Is CheckBox Then
               Call DrawCheckbox(DIS, bPushed, State, bEnabled, bFocus, bHover)
            ElseIf TypeOf Ctrl Is OptionButton Then
               Call DrawOption(DIS, bPushed, State, bEnabled, bFocus, bHover)
            Else
               Call DrawButton(DIS, bPushed, bEnabled, bFocus, bHover)
            End If
            FormProc = 1
         End If
      End If
   Else
      procOld = GetProp(hWnd, hWnd)
      FormProc = CallWindowProc(procOld, hWnd, uMsg, wParam, ByVal lParam)
   End If
End Function

Private Function ControlProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim procOld As Long
   Dim tp As POINTAPI
   Dim rc As RECT
   Dim bHover As Boolean
   
   procOld = GetProp(hWnd, hWnd)
   ControlProc = CallWindowProc(procOld, hWnd, uMsg, wParam, ByVal lParam)
   Select Case uMsg
   Case WM_MOUSEMOVE
      bHover = GetProp(hWnd, "Hover")
      If Not bHover Then
         GetCursorPos tp
         GetWindowRect hWnd, rc
         If Not (PtInRect(rc, tp.x, tp.y) = 0) Then
            SetProp hWnd, "Hover", -1
            SetTimer hWnd, 1, 10, 0
            InvalidateRectAsNull hWnd, 0, 0
            UpdateWindow hWnd
         End If
      End If
   Case WM_TIMER
      GetCursorPos tp
      GetWindowRect hWnd, rc
      If PtInRect(rc, tp.x, tp.y) = 0 Then
         KillTimer hWnd, 1
         RemoveProp hWnd, "Hover"
         InvalidateRectAsNull hWnd, 0, 0
         UpdateWindow hWnd
      End If
   End Select
End Function
