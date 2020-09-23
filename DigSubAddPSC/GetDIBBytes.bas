Attribute VB_Name = "GetDIBBytes"
' GetDIBBytes.bas  By  Robert Rayment

Option Explicit
Option Base 1


Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

'GetObjectAPI

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public PICWH As BITMAP


Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

'Public Type RGBQUAD
'        rgbBlue As Byte
'        rgbGreen As Byte
'        rgbRed As Byte
'        rgbReserved As Byte
'End Type

Public Type BITMAPINFO
   bmi As BITMAPINFOHEADER
   'Colors(0 To 255) As RGBQUAD
End Type


Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  system colors

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal x As Long, ByVal y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'StretchDIBits PICD.hDC, 0&, 0&, W4, H4, 0&, 0&, W, H, b8(1, 1), BS, DIB_RGB_COLORS, vbSrcCopy


' -----------------------------------------------------------

Private Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long
'----------------------------------------------------------------

Public Sub GETLONGS(ByVal PICINP As Long, _
   LA() As Long, bWidth As Long, bHeight As Long)
   
Dim BS As BITMAPINFO
Dim NewDC As Long
Dim OldH As Long

   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PICINP)
   With BS.bmi
      .biSize = 40
      .biwidth = bWidth
      .biheight = -bHeight
      .biPlanes = 1
      .biBitCount = 32     ' 32-bit colors
      .biCompression = 0
      .biSizeImage = 4 * bWidth * Abs(bHeight)
   End With
   
   If GetDIBits(NewDC, PICINP, 0, bHeight, LA(1, 1), BS, DIB_PAL_COLORS) = 0 Then
      MsgBox "DIB Error in GETLONGS 32bpp", vbCritical, " "
      End
   End If
   
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC
End Sub
