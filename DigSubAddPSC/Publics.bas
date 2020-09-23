Attribute VB_Name = "Publics"
Option Explicit


''------------------------------------------------------------------------------
''  This required instead of Screen.Height & Width for resizing
Public Declare Function GetSystemMetrics Lib "User32" _
   (ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0  ' Screen Width
Public Const SM_CYSCREEN = 1  ' Screen Height
Public Const SM_CYCAPTION = 4 ' Height of window caption
Public Const SM_CYMENU = 15   ' Height of menu
Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)
Public ExtraBorder As Long, ExtraHeight As Long
''------------------------------------------------------------


' Variables:
' PicBox, Array,    Width,   Height,  Loaded boolean
' PIC(0), ARR0(),   PICW(0), PICH(0), aPIC(0),        picture 0
' PIC(1), ARR1(),   PICW(1), PICH(1), aPIC(1),        picture 1
' PICRes, ARRREs()  resulting picture

' Weighting, GreyLevel, TheMode & InvertYN


Public ARR0() As Long
Public ARR1() As Long
Public ARRT() As Long   ' Temp for Saving Result
Public ARRRes() As Long
Public PICW() As Long
Public PICH() As Long
Public Weighting As Long
Public GreyLevel As Long
Public InvertYN As Long
Public TheMode As Long
Public UX As Long, UY As Long ' UBounds(ARRRes())
Public ix1 As Long, iy1 As Long, ix2 As Long, iy2 As Long ' ARRRES() on ARR0() RECT

Public aPIC() As Boolean
Public aSBarsActive As Boolean
Public HSMax() As Long, HSMin() As Long
Public VSMax() As Long, VSMin() As Long
Public sorcX() As Long, sorcY() As Long
Public HSResMax As Long, HSResMin As Long
Public VSResMax As Long, VSResMin As Long
Public sorcXRes As Long, sorcYRes As Long
Public aMouseDown As Boolean

Public picSmallX As Long, picSmallY As Long



Public PathSpec$, CurrPath$, FileSpec$
Public InFileSpec$()

Public Const PICWOrg As Long = 256
Public Const PICHOrg As Long = 256
Public PICResWOrg As Long '= 620
Public PICResHOrg As Long '= 500
Public PICResWDef As Long '= 620
Public PICResHDef As Long '= 500
Public Const FormWOrg As Long = 14200
Public Const FormHOrg As Long = 9975

Public aSelect() As Boolean
Public AF As Long ' AF 0 - 128

Public AlphaFactor As Long
Public zAlpha As Single

Public STX As Long, STY As Long 'Twips/Pixel

Public frmDisplayLeft As Long
Public frmDisplayTop As Long

Public Mul As Long   ' +1 or -1

Public Sub LngToRGB(LCul As Long, R As Byte, G As Byte, B As Byte)
'Convert Long Colors() to RGB components
'IN:  LCUL
'OUT: R,G & B bytes
R = (LCul And &HFF&)
G = (LCul And &HFF00&) / &H100&
B = (LCul And &HFF0000) / &H10000
End Sub

Public Function FName$(FSpec$)
' VB5 compatible
Dim p1 As Long, p2 As Long
   If Len(FSpec$) < 2 Then
      FName$ = ""
      Exit Function
   End If
   p1 = 1
   Do
      p2 = InStr(p1, FSpec$, "\")
      If p2 = 0 Then Exit Do
      p1 = p2 + 1
   Loop
   If p1 = Len(FSpec$) Then
      FName$ = ""
      Exit Function
   End If
   FName$ = " " & Mid$(FSpec$, p1) & " "
'   If Len(FName$) > 20 Then
   If Len(FName$) > 35 Then
      FName$ = Left$(FName$, 8) & "..." & Right$(FName$, 8)
   End If
End Function

Public Function TheExt$(FSpec$)
   TheExt$ = InStr(1, FName$(FSpec$), ".")
   TheExt$ = LCase$(Right$(TheExt$, 3))
End Function

Public Sub GetExtras(BStyle As Byte)
' IN:  BStyle = Me.BorderStyle
' OUT: Public ExtraBorder, ExtraHeight

''------------------------------------------------------------------------------
''  This required instead of Screen.Height & Width for resizing
'Public Declare Function GetSystemMetrics Lib "user32" _
'(ByVal nIndex As Long) As Long
'
'Public Const SM_CXSCREEN = 0  ' Screen Width
'Public Const SM_CYSCREEN = 1  ' Screen Height
'Public Const SM_CYCAPTION = 4 ' Height of window caption
'Public Const SM_CYMENU = 15   ' Height of menu
'Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
'Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)
'Public ExtraBorder, ExtraHeight
''------------------------------------------------------------
Dim Border As Long
Dim CapHeight As Long
Dim MenuHeight As Long
' BStyle 1 to 5 (not 0)
' BStyle = Form1.BorderStyle

Border = GetSystemMetrics(SM_CXDLGFRAME)
If BStyle = 2 Or BStyle = 5 Then Border = Border + 1 ' Sizable
If BStyle > 3 Then
   CapHeight = GetSystemMetrics(SM_CYSMCAPTION) ' Small cap - ToolWindow
Else
   CapHeight = GetSystemMetrics(SM_CYCAPTION)   ' Standard cap
End If
ExtraBorder = 2 * Border
ExtraHeight = CapHeight + ExtraBorder

MenuHeight = GetSystemMetrics(SM_CYMENU)
ExtraHeight = CapHeight + MenuHeight + ExtraBorder

' Win98  ExtraBorder=6 or 8, ExtraHeight= 41 - 46
' WinXP  ExtraBorder=6 or 8, ExtraHeight= 44 - 54
End Sub

Public Sub GetPublicCoords(frm As Form)
' Public UX As Long, UY As Long ' UBounds(ARRRes())
' Public ix1 As Long, iy1 As Long, ix2 As Long, iy2 As Long ' ARRRES() on ARR0() RECT
   UX = UBound(ARRRes(), 1)
   UY = UBound(ARRRes(), 2)
   ix1 = frm.picSmall.Left + sorcXRes + 1
   If ix1 < 1 Then ix1 = 1
   ix2 = ix1 + PICW(1) - 1
   If ix2 > UBound(ARR0(), 1) Then ix2 = UBound(ARR0(), 1)
   iy2 = PICH(0) - frm.PICRes.Height + frm.picSmall.Top + PICH(1) - sorcYRes '+ 1
   If iy2 > UBound(ARR0(), 2) Then iy2 = UBound(ARR0(), 2)
   iy1 = iy2 - PICH(1) + 1
   If iy1 < 1 Then iy1 = 1
End Sub


