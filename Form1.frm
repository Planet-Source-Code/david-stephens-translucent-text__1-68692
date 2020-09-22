VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8700
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar OpacityScroll 
      Height          =   375
      LargeChange     =   10
      Left            =   2520
      Max             =   255
      TabIndex        =   1
      Top             =   4200
      Width           =   3735
   End
   Begin VB.PictureBox FadeAnswerPictureBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00A4E3EC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2640
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private Const AC_SRC_OVER = &H0

'Private Const pixR As Integer = 3
'Private Const pixG As Integer = 2
'Private Const pixB As Integer = 1

Const DT_BOTTOM = &H8
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_RIGHT = &H2
Const DT_TOP = &H0
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10

Const StringToPrint = "Hello There"

Private Type BitmapInfoHEADER '40 bytes
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
    rgbReserved As Byte
End Type

Private Type BitmapInfo
    Header As BitmapInfoHEADER
    Colors As RGBQUAD
End Type

Dim Pixels() As Byte
Dim BackgroundBitmap As BitmapInfo

Dim BF            As BlendFunction
Dim lBF           As Long
Dim ThisRectangle As RECT
Dim Str           As String
Dim BackGroundDC  As Long
Dim iBitmap       As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BlendFunction ' This structure holds the arguments required by Alphablend function to work
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long) 'Conver to long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BitmapInfo, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BitmapInfo, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Sub CopyBackGroundIntoPictureBox()

Dim ThisWidth   As Integer
Dim ThisHeight  As Integer
Dim XCoord      As Integer
Dim YCoord      As Integer

XCoord = FadeAnswerPictureBox.Left
YCoord = FadeAnswerPictureBox.Top

ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight

FadeAnswerPictureBox.Visible = False
BitBlt FadeAnswerPictureBox.hdc, 0, 0, ThisWidth, ThisHeight, Form1.hdc, XCoord, YCoord, vbSrcCopy 'The part of form1 behind the picturebox
FadeAnswerPictureBox.Visible = True

End Sub
Sub PrintTranslucentText(ByVal ThisText As String, ThisOpacity As Integer)

CopyBackgroundFromMemory
FadeAnswerPictureBox.ForeColor = RGB(129, 0, 0)
DrawText FadeAnswerPictureBox.hdc, StringToPrint, Len(StringToPrint), ThisRectangle, DT_WORDBREAK   ' Print text
AlphaBlendWithBackground (ThisOpacity)
FadeAnswerPictureBox.Refresh

End Sub
Sub FadeIn()

Dim Opacity As Integer

For Opacity = 0 To 160
    PrintTranslucentText StringToPrint, Opacity
    DoEvents: Sleep (1) ' Wait
Next Opacity

End Sub
Sub FadeOut()

Dim Opacity As Integer

For Opacity = 160 To 0 Step -1
    PrintTranslucentText StringToPrint, Opacity
    DoEvents: Sleep (1) ' Wait
Next Opacity

DoEvents: Sleep (2000)

End Sub
Private Sub Form_Activate()

Dim Opacity As Integer

SetRect ThisRectangle, 0, 0, FadeAnswerPictureBox.ScaleWidth, FadeAnswerPictureBox.ScaleHeight ' Set coordinates
FadeAnswerPictureBox.FontSize = 48
Form1.Refresh

CopyBackGroundIntoPictureBox 'The part of form 1 behind the picture box goes into the picture box
CopyBackgroundToMemory 'The picture box (part of form1) goes into memory to be used in Alphablending

Opacity = 127
OpacityScroll.Value = 127
PrintTranslucentText StringToPrint, Opacity

End Sub
Sub AlphaBlendWithBackground(ByVal BlendValue As Integer)

Dim ThisWidth   As Integer
Dim ThisHeight  As Integer

BF.BlendOp = AC_SRC_OVER
BF.BlendFlags = 0
BF.SourceConstantAlpha = 255 - BlendValue
BF.AlphaFormat = 0
    
RtlMoveMemory lBF, BF, 4 'Convert the BLENDFUNCTION-structure to a Long
 
ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight

AlphaBlend FadeAnswerPictureBox.hdc, 0, 0, ThisWidth, ThisHeight, BackGroundDC, 0, 0, ThisWidth, ThisHeight, lBF

End Sub
Sub CopyBackgroundFromMemory()

SetDIBits FadeAnswerPictureBox.hdc, FadeAnswerPictureBox.Image, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS
FadeAnswerPictureBox.Picture = FadeAnswerPictureBox.Image
    
End Sub
Sub CopyBackgroundToMemory()

Dim ThisWidth   As Integer
Dim ThisHeight  As Integer
Dim XCoord      As Integer
Dim YCoord      As Integer
Dim Bytes_per_scanLine As Integer
Dim x, y As Integer

XCoord = FadeAnswerPictureBox.Left
YCoord = FadeAnswerPictureBox.Top

ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight

With BackgroundBitmap.Header ' Prepare the bitmap description.
    .biSize = 40
    .biWidth = ThisWidth
    .biHeight = -ThisHeight 'Use negative height to scan top-down.
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    Bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    .biSizeImage = Bytes_per_scanLine * Abs(.biHeight)
End With

ReDim Pixels(1 To 4, 1 To FadeAnswerPictureBox.ScaleWidth, 1 To FadeAnswerPictureBox.ScaleHeight) 'Load the bitmap's data.

BackGroundDC = CreateCompatibleDC(0) 'Create a context
iBitmap = CreateDIBSection(BackGroundDC, BackgroundBitmap, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&) 'Create a blank picture on the BackBmp standards (W,H,bitdebth)
SelectObject BackGroundDC, iBitmap 'Copy the picture into the context to make the context useable just like a picturebox

GetDIBits FadeAnswerPictureBox.hdc, FadeAnswerPictureBox.Image, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS
SetDIBits BackGroundDC, iBitmap, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS

End Sub

Private Sub Form_Unload(Cancel As Integer)

DeleteObject iBitmap
DeleteDC BackGroundDC
  
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub OpacityScroll_Change()

Dim Opacity As Integer

Opacity = OpacityScroll.Value
PrintTranslucentText StringToPrint, Opacity

End Sub

Private Sub OpacityScroll_Scroll()
Dim Opacity As Integer

Opacity = OpacityScroll.Value
PrintTranslucentText StringToPrint, Opacity

End Sub
