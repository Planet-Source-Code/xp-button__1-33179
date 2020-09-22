Attribute VB_Name = "mDrawItem"
Option Explicit

Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public IDB_BUTTON As Long
Public IDB_CHECKBOX As Long
Public Const IDB_RADIO = 103

Public Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION As Long = &H2000
Public Const IMAGE_BITMAP = 0&
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
' Timer function
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hWndItem As Long
  hDC As Long
  rcItem As RECT
  itemData As Long
End Type

Public Const WM_TIMER = &H113
Public Const WM_MOUSEMOVE = &H200
Public Const WM_DRAWITEM = &H2B
Public Const ODT_BUTTON = 4

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' Get Object
Public Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAPFILEHEADER
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbAlpha As Byte
End Type
Private Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors() As RGBQUAD
End Type
Private Const DIB_RGB_COLORS = 0
Private Const DEFAULT_PALETTE = 15
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Draw Text
Private Const DT_CALCRECT = &H400
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_EXPANDTABS = &H40
Private Const DT_NOCLIP = &H100
Private Const DT_EDITCONTROL = &H2000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORD_ELLIPSIS = &H40000
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const TRANSPARENT = 1
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

' Draw
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

Private Const WM_GETTEXT = &HD
' Color
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' Get State
Public Const BM_GETCHECK = &HF0&
Public Const BM_GETSTATE = &HF2&
' check state of a radio button or check box
Public Const BST_PUSHED = 4
Public Const BST_FOCUS = 8

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
' DC
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Function DrawCaption(DIS As DRAWITEMSTRUCT, rc As RECT, Optional bEnabled As Boolean, Optional ByVal Alignment As Long) As Boolean
   Dim r As Long
   Dim s As String
   
   s = String$(1024, 0)
   r = SendMessage(DIS.hWndItem, WM_GETTEXT, 1024, ByVal s)
   If r > 0 Then
      s = Left$(s, r)
      rc.Left = 16
      rc.Right = DIS.rcItem.Right
      DrawText DIS.hDC, s, -1, rc, Alignment Or DT_WORDBREAK Or DT_CALCRECT
   ' Make vertical alignment
      r = (DIS.rcItem.Bottom - rc.Bottom) \ 2
      rc.Top = rc.Top + r
      rc.Bottom = rc.Bottom + r
      If Not bEnabled Then
         SetTextColor DIS.hDC, &H92A1A1
      End If
      SetBkMode DIS.hDC, TRANSPARENT
      DrawText DIS.hDC, s, -1, rc, Alignment Or DT_WORDBREAK
      
      DrawCaption = True
   End If
End Function

Private Sub FillColor(hDC As Long, rc As RECT, Color As OLE_COLOR)
  Dim hBrush As Long
  
  OleTranslateColor Color, 0, Color
  hBrush = CreateSolidBrush(Color)
  FillRect hDC, rc, hBrush
  DeleteObject hBrush
End Sub

Public Sub DrawButton(DIS As DRAWITEMSTRUCT, bPushed As Boolean, bEnabled As Boolean, bFocus As Boolean, bHover As Boolean)
  Dim hMemDC  As Long
  Dim hBitmap As Long
  Dim hOldBitmap As Long
  Dim bmp As BITMAP
  Dim rc As RECT
  Dim s As String
  Dim r As Long
  Dim y As Integer
  Dim w As Integer
  Dim h As Integer
  
' Draw background
  hMemDC = CreateCompatibleDC(0)
  If hMemDC Then
      hBitmap = LoadImage(App.hInstance, CLng(IDB_BUTTON), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
      'hBitmap = LoadImage(0, App.Path & "\button.bmp", IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or LR_LOADFROMFILE)
      If hBitmap Then
         GetObject hBitmap, Len(bmp), bmp
         w = bmp.bmWidth
         h = bmp.bmHeight / 5
         If bPushed Then
            y = 2 * h
         ElseIf bHover Then
            y = h
         ElseIf bFocus Then
            y = 4 * h
         ElseIf bEnabled Then
            y = 0
         Else
            y = 3 * h
         End If
         SelectObject hMemDC, hBitmap
         If bmp.bmBitsPixel = 32 Then
            ApplyAlphaBlend hMemDC, hBitmap, bmp, y, y + h, vbButtonFace
         End If
      ' Draw Left-Top, Right-Top, Right-Bottom, Left-Bottom
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, 4, 4, hMemDC, 1, y + 1, vbSrcCopy
        BitBlt DIS.hDC, DIS.rcItem.Right - 4, DIS.rcItem.Top, 4, 4, hMemDC, w - 5, y + 1, vbSrcCopy
        BitBlt DIS.hDC, DIS.rcItem.Right - 4, DIS.rcItem.Bottom - 4, 4, 4, hMemDC, w - 5, y + h - 5, vbSrcCopy
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Bottom - 4, 4, 4, hMemDC, 1, y + h - 5, vbSrcCopy
      ' Draw Left, Top, Right, Bottom
        StretchBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top + 4, 4, DIS.rcItem.Bottom - DIS.rcItem.Top - 8, hMemDC, 1, y + 5, 4, h - 10, vbSrcCopy
        StretchBlt DIS.hDC, DIS.rcItem.Left + 4, DIS.rcItem.Top, DIS.rcItem.Right - DIS.rcItem.Left - 8, 4, hMemDC, 5, y + 1, w - 10, 4, vbSrcCopy
        StretchBlt DIS.hDC, DIS.rcItem.Right - DIS.rcItem.Left - 4, DIS.rcItem.Top + 4, 4, DIS.rcItem.Bottom - DIS.rcItem.Top - 8, hMemDC, w - 5, y + 5, 4, h - 10, vbSrcCopy
        StretchBlt DIS.hDC, DIS.rcItem.Left + 4, DIS.rcItem.Bottom - DIS.rcItem.Top - 4, DIS.rcItem.Right - DIS.rcItem.Left - 8, 4, hMemDC, 5, y + h - 5, w - 10, 4, vbSrcCopy
      ' Draw surface
        StretchBlt DIS.hDC, DIS.rcItem.Left + 4, DIS.rcItem.Top + 4, DIS.rcItem.Right - DIS.rcItem.Left - 8, DIS.rcItem.Bottom - DIS.rcItem.Top - 8, hMemDC, 5, y + 5, w - 10, h - 10, vbSrcCopy
         
         DeleteObject SelectObject(hMemDC, hOldBitmap)
      End If
      DeleteDC hMemDC
   End If

   DrawCaption DIS, rc, bEnabled, DT_CENTER
End Sub

Public Sub DrawCheckbox(DIS As DRAWITEMSTRUCT, bPushed As Boolean, ByVal bChecked As Integer, bEnabled As Boolean, bFocus As Boolean, bHover As Boolean)
  Dim hMemDC  As Long
  Dim hBitmap As Long
  Dim hOldBitmap As Long
  Dim bmp As BITMAP
  Dim rc As RECT
  Dim s As String
  Dim r As Long
  Dim y As Integer
  Dim w As Integer
  Dim h As Integer
  
' Draw background
   FillColor DIS.hDC, DIS.rcItem, GetBkColor(DIS.hDC)
' Draw Picture
   hMemDC = CreateCompatibleDC(0)
   If hMemDC Then
      hBitmap = LoadImage(App.hInstance, CLng(IDB_CHECKBOX), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
      'hBitmap = LoadImage(0, App.Path & "\checkbox.bmp", IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or LR_LOADFROMFILE)
      If hBitmap Then
         GetObject hBitmap, Len(bmp), bmp
         w = bmp.bmWidth
         h = bmp.bmHeight / 12
         If bChecked = 1 Then
           y = 4 * h
         ElseIf bChecked = 2 Then
           y = 8 * h
         End If
         If bPushed Then
           y = y + 2 * h
         ElseIf bHover Then
           y = y + h
         ElseIf Not bEnabled Then
           y = y + 3 * h
         End If
         hOldBitmap = SelectObject(hMemDC, hBitmap)
         If bmp.bmBitsPixel = 32 Then
            ApplyAlphaBlend hMemDC, hBitmap, bmp, y, y + h, GetBkColor(DIS.hDC)
         End If
         BitBlt DIS.hDC, DIS.rcItem.Left, (DIS.rcItem.Bottom - h) \ 2, w, h, hMemDC, 0, y, vbSrcCopy
         DeleteObject SelectObject(hMemDC, hOldBitmap)
      End If
      DeleteDC hMemDC
   End If
' Draw Caption
  If DrawCaption(DIS, rc, bEnabled) Then
    If bFocus Then
      InflateRect rc, 1, 2
      rc.Bottom = rc.Bottom - 1
      If rc.Top < DIS.rcItem.Top Then
         rc.Top = DIS.rcItem.Top
      End If
      If rc.Bottom > DIS.rcItem.Bottom Then
         rc.Bottom = DIS.rcItem.Bottom
      End If
      DrawFocusRect DIS.hDC, rc
    End If
  End If
End Sub

Public Sub DrawOption(DIS As DRAWITEMSTRUCT, bPushed As Boolean, ByVal bChecked As Integer, bEnabled As Boolean, bFocus As Boolean, bHover As Boolean)
  Dim hMemDC  As Long
  Dim hBitmap As Long
  Dim hOldBitmap As Long
  Dim bmp As BITMAP
  Dim rc As RECT
  Dim s As String
  Dim r As Long
  Dim y As Integer
  Dim w As Integer
  Dim h As Integer
  
' Draw background
   FillColor DIS.hDC, DIS.rcItem, GetBkColor(DIS.hDC)
' Draw Picture
   hMemDC = CreateCompatibleDC(0)
   If hMemDC Then
      hBitmap = LoadImage(App.hInstance, CLng(IDB_RADIO), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
      'hBitmap = LoadImage(0, App.Path & "\radio.bmp", IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or LR_LOADFROMFILE)
      If hBitmap Then
         GetObject hBitmap, Len(bmp), bmp
         w = bmp.bmWidth
         h = bmp.bmHeight / 8
         If bChecked Then y = 4 * h
         If bPushed Then
            y = y + 2 * h
         ElseIf bHover Then
            y = y + h
         ElseIf Not bEnabled Then
            y = y + 3 * h
         End If
         hOldBitmap = SelectObject(hMemDC, hBitmap)
         If bmp.bmBitsPixel = 32 Then
            ApplyAlphaBlend hMemDC, hBitmap, bmp, y, y + h, GetBkColor(DIS.hDC)
         End If
         BitBlt DIS.hDC, DIS.rcItem.Left, (DIS.rcItem.Bottom - h) \ 2, w, h, hMemDC, 0, y, vbSrcCopy
         DeleteObject SelectObject(hMemDC, hOldBitmap)
      End If
      DeleteDC hMemDC
  End If
' Draw Caption
  If DrawCaption(DIS, rc, bEnabled) Then
    If bFocus Then
      InflateRect rc, 1, 1
      rc.Bottom = rc.Bottom + 1
      If rc.Top < DIS.rcItem.Top Then rc.Top = DIS.rcItem.Top
      If rc.Bottom > DIS.rcItem.Bottom Then rc.Bottom = DIS.rcItem.Bottom
      DrawFocusRect DIS.hDC, rc
    End If
  End If
End Sub

Private Sub ApplyAlphaBlend(hMemDC As Long, hBitmap As Long, bmp As BITMAP, Y1 As Integer, Y2 As Integer, bgColor As OLE_COLOR)
   On Error GoTo ErrorHandler
   Dim bmi As BITMAPINFO
   Dim BytesPerScanLine As Long
   Dim Buffer() As Byte
   
   Dim a As Long
   Dim r As Long
   Dim g As Long
   Dim b As Long
   
   Dim x As Integer
   Dim y As Integer
   
   With bmi.bmiHeader
      .biSize = Len(bmi.bmiHeader)
      .biWidth = bmp.bmWidth
      .biHeight = -bmp.bmHeight
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = 0
      BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
      
      .biSizeImage = BytesPerScanLine * bmp.bmHeight
      ReDim Buffer(3, .biWidth - 1, bmp.bmHeight - 1)
      GetDIBits hMemDC, hBitmap, 0, bmp.bmHeight, Buffer(0, 0, 0), bmi, DIB_RGB_COLORS
   End With
   
   ExtractColor bgColor, r, g, b
   For y = Y1 To Y2 - 1
      For x = 0 To bmi.bmiHeader.biWidth - 1
         a = Buffer(3, x, y)
'         Buffer(0, x, y) = a
'         Buffer(1, x, y) = a
'         Buffer(2, x, y) = a
         If a = 0 Then
            Buffer(0, x, y) = b
            Buffer(1, x, y) = g
            Buffer(2, x, y) = r
         ElseIf a = 255 Then
            Buffer(0, x, y) = Buffer(0, x, y)
            Buffer(1, x, y) = Buffer(1, x, y)
            Buffer(2, x, y) = Buffer(2, x, y)
         Else
            Buffer(0, x, y) = (b * (255 - a) + a * Buffer(0, x, y)) \ 255
            Buffer(1, x, y) = (g * (255 - a) + a * Buffer(1, x, y)) \ 255
            Buffer(2, x, y) = (r * (255 - a) + a * Buffer(2, x, y)) \ 255
         End If
      Next
   Next
   SetDIBits hMemDC, hBitmap, 0, bmp.bmHeight, Buffer(0, 0, 0), bmi, DIB_RGB_COLORS
   Exit Sub
   
ErrorHandler:
   Debug.Print x, y, Y1, Y2, bmp.bmWidth - 1, bmp.bmHeight - 1
End Sub

Private Sub ExtractColor(ByVal Color As Long, r As Long, g As Long, b As Long)
   OleTranslateColor Color, 0, Color
   b = Color \ &H10000 And &HFF
   g = Color \ &H100 And &HFF
   r = Color And &HFF
End Sub
