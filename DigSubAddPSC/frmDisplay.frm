VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   90
      Width           =   1290
   End
   Begin VB.PictureBox PICD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   180
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   465
      Width           =   750
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDisplay.frm

Option Explicit

Private Sub cmdClose_Click()
   frmDisplayLeft = Me.Left
   frmDisplayTop = Me.Top
   PICD.Picture = LoadPicture
   PICD.Width = 6
   PICD.Height = 6

   Unload Me
End Sub

Private Sub Form_Load()
Dim ix As Long, iy As Long, ixx As Long, iyy As Long
Dim W As Long, H As Long
Dim WD As Long, HD As Long
Dim UXX As Long, UYY As Long
Dim BS As BITMAPINFO
   ' PICD default= 600x400
   WD = 600
   HD = 400
   PICD.Width = WD
   PICD.Height = HD
   W = PICW(0)
   H = PICH(0)
   ReDim ARRT(1 To W, 1 To H)
   ARRT() = ARR0()
   ' Copy ARRRes() rectangle to ARRT()
   GetPublicCoords Form1
   iyy = 1
   For iy = iy1 To iy2
      ixx = 1
      For ix = ix1 To ix2
         If ixx <= UX Then
         If iyy <= UY Then
            ARRT(ix, iy) = ARRRes(ixx, iyy)
        End If
        End If
        ixx = ixx + 1
      Next ix
      iyy = iyy + 1
   Next iy
   
   ' Size to PICD
   If W <= WD And H <= HD Then
      PICD.Width = W
      PICD.Height = H
      WD = PICD.Width
      HD = PICD.Height
   ElseIf W >= H Then   ' WD=600
      HD = H * (WD / W)
      PICD.Height = HD
   ElseIf H > W Then    ' HD=400
      WD = W * (HD / H)
      PICD.Width = WD
   End If
   
   Me.Width = PICD.Width * STX + 450
   Me.Height = PICD.Height * STY + 900 + 300
   Me.Left = frmDisplayLeft
   Me.Top = frmDisplayTop

   
   ' Display
   UXX = UBound(ARRT(), 1)
   UYY = UBound(ARRT(), 2)
   With BS.bmi
      .biSize = 40
      .biwidth = UXX
      .biheight = UYY * -1 'Mul ' if -1 required
      .biPlanes = 1
      .biBitCount = 32    ' Sets up 32-bit colors
      .biCompression = 0
      .biSizeImage = 0 'UXX * UYY * 4
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   PICD.Picture = LoadPicture
      
   SetStretchBltMode PICD.hdc, 4  ' NB Of Dest picbox
   If StretchDIBits(PICD.hdc, 0, 0, WD, HD, 0, 0, _
      W - 1, H - 1, ARRT(1, 1), BS, DIB_PAL_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Display StretchDIBits Error", vbCritical, " "
   End If
   PICD.Refresh
   Erase ARRT()
End Sub
