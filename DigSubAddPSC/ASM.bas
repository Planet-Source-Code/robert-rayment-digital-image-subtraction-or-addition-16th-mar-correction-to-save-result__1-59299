Attribute VB_Name = "ASM"
'ASM.bas

Option Explicit


Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Single, _
ByVal Long3 As Single, ByVal Long4 As Long) As Long

Public MMXCode() As Byte     'Array to hold machine code

'MCode Structure
Public Type MCodeStruc
   PICW0 As Long
   PtrARR0 As Long
   PICW1 As Long
   PICH1 As Long
   PtrARR1 As Long
   PtrARRRES As Long
   MODE As Long         'Subtraction Mode 0,1 (Kind of Subtraction 0=normal(Mode 1), 1=xor(Mode 2) , 2=.....(Mode 3) ,.....)
                        'Add Modes 2,3 Alpha, Edge Alpha
   BGL As Long          'Base Grey Level
   WDM As Long          'Weighting Factor * DiffMul
   UX As Long
   UY As Long
   ix1 As Long
   ix2 As Long
   iy1 As Long
   iy2 As Long
   ALPH As Long
End Type
Public MCode As MCodeStruc
'-------------------------------------


Public Sub ASM_DigSub(frm As Form)
Dim ptrStruc As Long
Dim ptrMC As Long
Dim res As Long
Dim Index As Long
Dim DiffMul As Long
Dim WDM As Long

   GetPublicCoords frm
   
   If aSelect(0) Then   'Subtract
      MCode.MODE = TheMode   ' 0,1
   Else  ' Add
      MCode.MODE = AlphaFactor + 2   ' 0+2,1+2 = 2,3
   End If
   
   DiffMul = 1
   If InvertYN = 1 Then DiffMul = -1
   WDM = Weighting * DiffMul
   
   If MCode.MODE = 1 Then WDM = -WDM
   
   ' Fill MCodeStruc
   With MCode
      .PICW0 = PICW(0)
      .PtrARR0 = VarPtr(ARR0(1, 1))
      .PICW1 = PICW(1)
      .PICH1 = PICH(1)
      .PtrARR1 = VarPtr(ARR1(1, 1))
      .PtrARRRES = VarPtr(ARRRes(1, 1))
      .BGL = GreyLevel
      .WDM = WDM '(Weighting * DiffMul)
      .UX = UX
      .UY = UY
      .ix1 = ix1
      .ix2 = ix2
      .iy1 = iy1
      .iy2 = iy2
      .ALPH = AF
   End With
   ptrStruc = VarPtr(MCode.PICW0)
   ptrMC = VarPtr(MMXCode(0))
   res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)
End Sub

