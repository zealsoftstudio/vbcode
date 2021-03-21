Attribute VB_Name = "Module1"
Option Explicit

Public Const GCL_HCURSOR = -12

Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Any) As Long
Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

