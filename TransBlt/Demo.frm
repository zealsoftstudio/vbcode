VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Transparent Copy Demo"
   ClientHeight    =   4092
   ClientLeft      =   888
   ClientTop       =   1428
   ClientWidth     =   7680
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4464
   Left            =   840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4092
   ScaleWidth      =   7680
   Top             =   1104
   Width           =   7776
   Begin VB.PictureBox Picture1 
      Height          =   1092
      Left            =   5880
      ScaleHeight     =   1044
      ScaleWidth      =   1404
      TabIndex        =   11
      Top             =   2280
      Width           =   1452
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Change &Destination picture"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Change &source picture"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.PictureBox pictSource 
      Height          =   2052
      Left            =   360
      Picture         =   "DEMO.frx":0000
      ScaleHeight     =   2004
      ScaleWidth      =   2484
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.PictureBox pictDest 
      Height          =   2055
      Left            =   3120
      Picture         =   "DEMO.frx":0282
      ScaleHeight     =   2004
      ScaleWidth      =   2604
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DefaultExt      =   "bmp"
      Filter          =   "Bitmap|*.bmp|All|*.*"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Click source picture to change transparent color."
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   4185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Destination Picture:"
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Source Picture:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2040
      Shape           =   1  'Square
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transparent Color:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Written by Hai Li, Zeal SoftStuido. http://www.nease.net/~zealsoft/indexc.html"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'
' Copyright (c) Hai Li, Zeal SoftStudio 1997
' All Rights Reserved.
' Email: haili@public.bta.net.cn
' http://www.nease.net/~zealsoft/indexc.html
'
' Demo of TransparentBlt function.
' May be freely used in your applications.
'
Dim cTransparent As Long
#If Win32 Then
    Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type
    Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal _
        hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
#Else
    Private Type BITMAP
        bmType As Integer
        bmWidth As Integer
        bmHeight As Integer
        bmWidthBytes As Integer
        bmPlanes As String * 1
        bmBitsPixel As String * 1
        bmBits As Long
    End Type
    Private Declare Function GetObj Lib "GDI" Alias "GetObject" (ByVal hObject _
        As Integer, ByVal nCount As Integer, bmp As Any) As Integer
#End If

Private Sub Command1_Click()
    Dim bmp As BITMAP
    
    ' Get the dimension of specific bitmap
    GetObj pictSource.Picture, Len(bmp), bmp
    TransparentBlt pictDest.hdc, pictSource.hdc, _
        0, 0, bmp.bmWidth, bmp.bmHeight, 0, 0, cTransparent
End Sub

Private Sub Command2_Click()
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
        pictSource.Picture = LoadPicture(CommonDialog1.filename)
    End If
End Sub

Private Sub Command3_Click()
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
        pictDest.Picture = LoadPicture(CommonDialog1.filename)
    End If
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub pictSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cTransparent = pictSource.Point(x, y)
    pictDest.Refresh
    Picture1.Refresh
    Shape1.FillColor = cTransparent
End Sub

Private Sub Form_Activate()
    cTransparent = pictSource.Point(0, 0)
    Shape1.FillColor = cTransparent
End Sub

