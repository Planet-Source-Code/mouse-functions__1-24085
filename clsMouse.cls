VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_MOVE = &H1

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Public Sub DisableTrap()

Dim erg As Long
Dim NewRect As RECT

  With NewRect
    .Left = 0&
    .Top = 0&
    .Right = Screen.Width / Screen.TwipsPerPixelX
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
  End With
  
  erg& = ClipCursor(NewRect)
  
End Sub

Public Sub EnableTrap(lHwnd As Long)

Dim ret As Long, erg As Long
Dim p As POINTAPI
Dim NewRect As RECT
  
  ret = GetWindowRect(lHwnd, NewRect)
  erg& = ClipCursor(NewRect)
  
End Sub

Public Sub LeftClick(lHwnd As Long, Optional x As Long, Optional y As Long)

'An easier way to click a VB command button is to set its value to true.
'Command1.Value = True

'x and y are the position withing the handle

Dim ret&
Dim p As POINTAPI

  p.x = 0
  p.y = 0
  
  ret = ClientToScreen(lHwnd, p)
  
  If x Then p.x = p.x + x
  If y Then p.y = p.y + y
  
  ret = SetCursorPos(p.x, p.y)
  mouse_event MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0

End Sub

Public Sub RightClick(lHwnd As Long, Optional x As Long, Optional y As Long)

'x and y are the position withing the handle

Dim ret&
Dim p As POINTAPI

  p.x = 0
  p.y = 0
  
  ret = ClientToScreen(lHwnd, p)
  
  If x Then p.x = p.x + x
  If y Then p.y = p.y + y
  
  ret = SetCursorPos(p.x, p.y)
  mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, p.x, p.y, 0, 0

End Sub

Public Sub CursorVisible(Optional TrueFalse As Boolean = True)

'Set TrueFalse to True to show the cursor, false to hide it

Dim ret As Long
  
  'Windows maintains an internal count that is incremented each time ShowCursor is called.
  'Without this loop, if you hide the cursor 5 times in a row, you would have to show the
  'cursor 5 times in a row to bring it back.
  
  If TrueFalse = False Then
    ret = ShowCursor(0)
    Do Until ret <= 0
      ret = ShowCursor(0)
    Loop
  Else
    ret = ShowCursor(1)
    Do Until ret > 0
      ret = ShowCursor(1)
    Loop
  End If
  
End Sub

Public Sub SwapButtons(Optional Normal As Boolean = True)

'Swap the left and right mouse buttons

Dim ret As Long

  If Normal = True Then
    ret = SwapMouseButton(0) 'Normal
  Else
    ret = SwapMouseButton(1) 'Switched
  End If

End Sub
