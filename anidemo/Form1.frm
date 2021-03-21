VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "AniDemo"
   ClientHeight    =   2220
   ClientLeft      =   912
   ClientTop       =   1332
   ClientWidth     =   3804
   Height          =   2592
   Left            =   864
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   3804
   Top             =   1008
   Width           =   3900
   Begin VB.TextBox Text1 
      Height          =   864
      Left            =   108
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   372
      Width           =   3540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Common"
      Height          =   396
      Left            =   2664
      TabIndex        =   1
      Top             =   1524
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ani Cursor"
      Height          =   396
      Left            =   1476
      TabIndex        =   0
      Top             =   1524
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim mhBaseCursor As Long, mhAniCursor As Long
Dim mhBaseCursor2 As Long, mhAniCursor2 As Long
Dim state As Integer
Private Sub Command1_Click()
    Dim lResult As Long
    
    mhAniCursor = LoadCursorFromFile("c:\win95\cursors\appstart.ani")
    lResult = SetClassLong((hwnd), GCL_HCURSOR, mhAniCursor)
    state = 1
    
    mhAniCursor2 = LoadCursorFromFile("c:\win95\cursors\Pen_1.cur")
    lResult = SetClassLong((Text1.hwnd), GCL_HCURSOR, mhAniCursor2)
    state = 1
End Sub

Private Sub Command2_Click()
    Dim lResult As Long
    
    lResult = SetClassLong((hwnd), GCL_HCURSOR, mhBaseCursor)
    lResult = DestroyCursor(mhAniCursor)
    
    lResult = SetClassLong((Text1.hwnd), GCL_HCURSOR, mhBaseCursor2)
    lResult = DestroyCursor(mhAniCursor2)
    state = 0
End Sub


Private Sub Form_Load()

    mhBaseCursor = GetClassLong((hwnd), GCL_HCURSOR)
    mhBaseCursor2 = GetClassLong((hwnd), GCL_HCURSOR)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If state Then Command2_Click
End Sub


