VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Digital Image Subtraction or Addition"
   ClientHeight    =   9840
   ClientLeft      =   165
   ClientTop       =   -825
   ClientWidth     =   14085
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   656
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   939
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   480
      Left            =   12645
      TabIndex        =   36
      Top             =   360
      Width           =   660
   End
   Begin VB.PictureBox picL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   11250
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   32
      Top             =   60
      Width           =   1200
      Begin VB.PictureBox picS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   255
         ScaleHeight     =   38
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   38
         TabIndex        =   33
         Top             =   255
         Width           =   600
      End
   End
   Begin Project1.dmFrame Frame2 
      Height          =   1230
      Left            =   6525
      TabIndex        =   24
      Top             =   45
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   2170
      Caption         =   "Digital Addition"
      BarColor        =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdRun2 
         BackColor       =   &H80000013&
         Caption         =   "0,0"
         Height          =   270
         Index           =   2
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   " Locate & top-left "
         Top             =   300
         Width           =   570
      End
      Begin VB.CommandButton cmdRun2 
         Caption         =   "RunA"
         Height          =   270
         Index           =   1
         Left            =   3315
         TabIndex        =   31
         Top             =   300
         Width           =   570
      End
      Begin VB.HScrollBar HSAlpha 
         Height          =   240
         LargeChange     =   5
         Left            =   105
         Max             =   100
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   780
         Width           =   3765
      End
      Begin VB.CommandButton cmdRun2 
         Caption         =   "RunB"
         Height          =   270
         Index           =   0
         Left            =   2655
         TabIndex        =   27
         Top             =   300
         Width           =   570
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Alpha edges"
         Height          =   210
         Index           =   1
         Left            =   1335
         TabIndex        =   26
         Top             =   345
         Width           =   1215
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add alpha"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   345
         Width           =   1050
      End
      Begin VB.Label LabAF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AF"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3915
         TabIndex        =   29
         Top             =   780
         Width           =   420
      End
   End
   Begin Project1.dmFrame Frame1 
      Height          =   1230
      Left            =   5895
      TabIndex        =   12
      Top             =   45
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   2170
      Caption         =   "Digital Subtraction"
      BarColor        =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdRun 
         BackColor       =   &H80000013&
         Caption         =   "0,0"
         Height          =   270
         Index           =   2
         Left            =   3885
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   " Locate & top-left "
         Top             =   270
         Width           =   570
      End
      Begin VB.HScrollBar HSGreyLevel 
         Height          =   255
         Left            =   870
         TabIndex        =   22
         Top             =   885
         Width           =   2475
      End
      Begin VB.HScrollBar HSWeighting 
         Height          =   270
         Left            =   870
         TabIndex        =   19
         Top             =   585
         Width           =   2475
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "RunA"
         Height          =   270
         Index           =   1
         Left            =   3120
         TabIndex        =   17
         Top             =   270
         Width           =   630
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run B"
         Height          =   270
         Index           =   0
         Left            =   2370
         TabIndex        =   16
         Top             =   270
         Width           =   690
      End
      Begin VB.CheckBox chkInvert 
         Caption         =   "Invert"
         Height          =   225
         Left            =   1560
         TabIndex        =   15
         Top             =   285
         Width           =   750
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Xor"
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   14
         Top             =   270
         Width           =   555
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Minus"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Top             =   270
         Width           =   750
      End
      Begin VB.Label LabGreyLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3375
         TabIndex        =   23
         Top             =   870
         Width           =   465
      End
      Begin VB.Label LabG 
         Caption         =   "GreyLevel"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   870
         Width           =   765
      End
      Begin VB.Label LabWeighting 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3375
         TabIndex        =   20
         Top             =   600
         Width           =   465
      End
      Begin VB.Label LabW 
         Caption         =   "Weighting"
         Height          =   210
         Left            =   75
         TabIndex        =   18
         Top             =   585
         Width           =   840
      End
   End
   Begin Project1.dmFrame Frame3 
      Height          =   1230
      Left            =   4545
      TabIndex        =   9
      Top             =   45
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   2170
      Caption         =   "Select"
      BarColor        =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optSelect 
         Caption         =   "Add"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   810
         Width           =   750
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Subtract"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   435
         Width           =   1005
      End
   End
   Begin VB.PictureBox PIC_Save 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   13635
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   7
      Top             =   495
      Width           =   510
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Pic 1 <= Pic 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1290
      TabIndex        =   4
      Top             =   4485
      Width           =   1860
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Pic 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   90
      Width           =   1080
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Index           =   1
      Left            =   330
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   5130
      Width           =   3840
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Index           =   0
      Left            =   330
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   420
      Width           =   3840
   End
   Begin VB.PictureBox PICRes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   4590
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   614
      TabIndex        =   2
      Top             =   1425
      Width           =   9210
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   135
         MousePointer    =   5  'Size
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   82
         TabIndex        =   8
         Top             =   165
         Width           =   1230
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Digital MMX by Robert Rayment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   570
      TabIndex        =   34
      Top             =   9465
      Width           =   2955
   End
   Begin VB.Label LabNWH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   4815
      Width           =   60
   End
   Begin VB.Label LabNWH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1725
      TabIndex        =   5
      Top             =   105
      Width           =   60
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Save Result"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Digital Subtraction & Addition by  Robert Rayment

' 16/03/05
' Correction to Save Result

' ASM action Scrollbars & Smaller image movemebt after RunB or RunA
' HandScrollers only, no Scrollbars
' Modified dreamVB Frame XP added

' Formulae:
' Mode 0 Minus:  GreyLevel +/- (RGB0 - RGB1) * Weighting
' Mode 1 Xor:    GreyLevel -/+ (RGB0 XOR RGB1) * Weighting
'                Invert swaps sign

' Alpha:   varies transparency of pic 1 on picSmall         ' Mode 2
' Feather: varies edging transparency of pic 1 on picSmall  ' Mode 3

' RunB  Run BASIC routine
' RunA  Run ASM routine

' Main.frm Form1

Option Explicit

Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1

Private aRun As Boolean
Private aPic0EqPic1 As Boolean

Dim CommonDialog1 As OSDialog

Private Sub cmdDisplay_Click()
   If aSBarsActive And aRun Then
      frmDisplay.Show 1
   End If
End Sub

' Variables:
' PicBox,  Array,    Width,   Height,  Loaded boolean
' PIC(0),  ARR0(),   PICW(0), PICH(0), aPIC(0),        picture 0
' PIC(1),  ARR1(),   PICW(1), PICH(1), aPIC(1),        picture 1
' PICRes,  ARRREs()  resulting picture or
' picSmall ARRREs()  resulting picture
' Weighting & GreyLevel & InvertYN

' Public Const PICWOrg As Long = 256
' Public Const PICHOrg As Long = 256
' Public Const PICResWOrg As Long = 620
' Public Const PICResHOrg As Long = 500

Private Sub Form_Load()
'Public Const FormWOrg As Long = 14200
'Public Const FormHOrg As Long = 9975

   PIC_Save.Visible = False
   picSmall.Visible = False

   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY

   If Screen.Width \ STX < 1024 Then
      MsgBox "Minimum screen resolution must be >= 1024 x 768", vbCritical, "DigSub"
      Unload Me
      End
   End If
   
   Me.Width = FormWOrg
   Me.Height = FormHOrg
   PICResWDef = 620
   PICResHDef = 500
   PICResWOrg = PICResWDef
   PICResHOrg = PICResHDef
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrPath$ = PathSpec$

   ReDim HSMax(2), HSMin(2)
   ReDim VSMax(2), VSMin(2)
   ReDim sorcX(2), sorcY(2)
   ReDim InFileSpec$(2)
   
   PositionControls
   ReDim aPIC(2)
   ReDim PICW(2), PICH(2)
   ReDim aPIC(2)
   ' Start values
   aPIC(0) = False
   aPIC(1) = False
   TheMode = 0
   InvertYN = 0
   optMode(0).Value = True
   chkInvert.Value = Unchecked
   cmdRun(0).Enabled = False
   cmdRun(1).Enabled = False
   cmdDisplay.Enabled = False
   mnuFile(0).Enabled = False
   aRun = False
   
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'To Load Machine code frm Res file
' Public MMXCode() As Byte     'Array to hold machine code
 MMXCode = LoadResData("MMXMC", "CUSTOM")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   
   ReDim aSelect(2)
   aSelect(0) = False
   aSelect(1) = False
   optSelect(0).Value = False
   optSelect(1).Value = False
   optSelect(0).Enabled = False
   optSelect(1).Enabled = False
   
   AlphaFactor = 0
   optAdd(0).Value = True
   AF = 50
   zAlpha = 0.5
   HSAlpha.Value = AF
   LabAF = Str$(AF)
   
   Frame1.Visible = False   ' Left = 390
   Frame2.Visible = False
   
   PICRes.MousePointer = vbCustom
   PICRes.MouseIcon = LoadResPicture("OPENHAND", vbResCursor)
   PIC(0).MousePointer = vbCustom
   PIC(0).MouseIcon = LoadResPicture("SMALLOPENHAND", vbResCursor)
   PIC(1).MousePointer = vbCustom
   PIC(1).MouseIcon = LoadResPicture("SMALLOPENHAND", vbResCursor)
   
   frmDisplayLeft = 200
   frmDisplayTop = 200
End Sub

Private Sub cmdLoad_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim iBPP As Integer
Dim aTT As Boolean
   aRun = False
   aPIC(Index) = False
   mnuFile(0).Enabled = False
   
   Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
   'Filt$ = "BMP(*.bmp)|*.bmp|JPEG(*.jpg)|*.jpg|GIF(*.gif)|*.gif"

   FileSpec$ = ""
   Title$ = "Load PIC" & Str$(Index)
   InDir$ = CurrPath$ 'Pathspec$
   Set CommonDialog1 = New OSDialog
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
   Set CommonDialog1 = Nothing
   If Len(FileSpec$) > 0 Then
      CurrPath$ = FileSpec$
      
      PIC(Index).Picture = LoadPicture(FileSpec$)
      If TheExt$(FileSpec$) = "bmp" Then Mul = 1 Else Mul = -1
      'If FIndex = 1 Then Mul = 1 Else Mul = -1
      GetObjectAPI PIC(Index), Len(PICWH), PICWH
      iBPP = PICWH.bmBitsPixel      ' 24 bpp
      PICW(Index) = PICWH.bmWidth
      PICH(Index) = PICWH.bmHeight
      
      If Index = 0 Then
         InFileSpec$(0) = FileSpec$
         ReDim ARR0(1 To PICW(0), 1 To PICH(0))
         GETLONGS PIC(0).Picture, ARR0(), PICW(Index), PICH(Index)
         aPIC(0) = True
      Else
         InFileSpec$(1) = FileSpec$
         ReDim ARR1(1 To PICW(1), 1 To PICH(1))
         GETLONGS PIC(1).Picture, ARR1(), PICW(Index), PICH(Index)
         aPIC(1) = True
      End If
      DoEvents
      LabNWH(Index) = " " & FName$(FileSpec$) & Str$(PICW(Index)) & " x" & Str$(PICH(Index)) & " "
      PICRes.SetFocus
   End If

   aTT = ((PICW(0) < PICW(1)) Or _
         (PICH(0) < PICH(1)))
   
   If aPIC(0) And aPIC(1) Then
      If aTT Then
         MsgBox "Make sure Pic 1 width & height is <= Pic 0" & vbCrLf _
         & " Pic1 width & height is <" & Str$(PICWOrg) & " &" & Str$(PICHOrg), vbCritical, " "
         cmdRun(0).Enabled = False
         cmdRun(1).Enabled = False
         cmdRun2(0).Enabled = False
         cmdRun2(1).Enabled = False
         Exit Sub
      End If
      Mul = -1
      SetScrollBars     ' Also does SetResScrollBars
      
      cmdRun(0).Enabled = True
      cmdRun(1).Enabled = True  ' ASM
      cmdRun2(0).Enabled = True
      cmdRun2(1).Enabled = True
      optSelect(0).Enabled = True
      optSelect(1).Enabled = True
      picSmall.Left = 0
      picSmall.Top = 0
      aPic0EqPic1 = ((PICW(0) = PICW(1)) And (PICH(0) = PICH(1)))
      
      If PICW(0) >= PICH(0) Then
         picL.Width = 80
         picL.Height = 80 * (PICH(0) / PICW(0))
      Else
         picL.Height = 80
         picL.Width = 80 * (PICW(0) / PICH(0))
      End If
      picS.Width = picL.Width * (PICW(1) / PICW(0))
      picS.Height = picL.Height * (PICH(1) / PICH(0))
      picS.Left = 0
      picS.Top = 0
      
      DisplaySmallPic
      
      PICRes.Cls
      picSmall.Cls
   End If
End Sub

Private Sub DisplaySmallPic()
Dim UX As Long, UY As Long
Dim BS As BITMAPINFO
   UX = UBound(ARR0(), 1)
   UY = UBound(ARR0(), 2)
   With BS.bmi
      .biSize = 40
      .biwidth = UX
      .biheight = UY * -1 'Mul ' if -1 required
      .biPlanes = 1
      .biBitCount = 32    ' Sets up 32-bit colors
      .biCompression = 0
      .biSizeImage = 0 'UX * UY * 4
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   picL.Picture = LoadPicture
   SetStretchBltMode picL.hdc, 4  ' NB Of Dest picbox
   If StretchDIBits(picL.hdc, 0, 0, picL.Width, picL.Height, 0, 0, _
      UX, UY, ARR0(1, 1), BS, DIB_PAL_COLORS, vbSrcCopy) = 0 Then
         MsgBox "StretchDIBits Error", vbCritical, " "
   End If
   picL.Refresh
End Sub

Private Sub cmdRun_Click(Index As Integer)
'Public Const PICResWOrg As Long = 620
'Public Const PICResHOrg As Long = 480
Dim sx1 As Long
Dim sy1 As Long, sy2 As Long
   ' To Set PICRes size as image
   If PICW(0) <= PICRes.Width Then
      PICRes.Width = PICW(0)
   Else
      If PICW(0) <= PICResWOrg Then
         PICRes.Width = PICW(0)
      Else
         PICRes.Width = PICResWOrg
      End If
   End If
   If PICH(0) <= PICRes.Height Then
      PICRes.Height = PICH(0)
   Else
      If PICH(0) <= PICResHOrg Then
         PICRes.Height = PICH(0)
      Else
         PICRes.Height = PICResHOrg
      End If
   End If
   PICRes.Picture = LoadPicture  ' To Flash
   PICRes.Refresh
   
   If aRun = False Then
      SetScrollBars     ' Also does SetResScrollBars
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARR0(), sorcXRes, sorcYRes, -1
   Else
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARR0(), sorcXRes, sorcYRes, -1
   End If
      
   Select Case Index
   Case 0   ' RunB
      picSmall.Visible = True
      picSmall.Picture = LoadPicture(InFileSpec$(1))
      ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
      RunBASIC
      DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   Case 1   ' RunA
      picSmall.Visible = True
      picSmall.Picture = LoadPicture(InFileSpec$(1))
      ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
      ASM_DigSub Me
      DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   Case 2   ' 0,0
      picSmall.Left = 0
      picSmall.Top = 0
      If aSBarsActive And aRun Then
         picSmall.Visible = True
         picSmall.Picture = LoadPicture(InFileSpec$(1))
         ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
         ARRRes() = ARR1()
         ASM_DigSub Me
         DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
      End If
      ' Small view
      sx1 = picSmall.Left + sorcXRes
      sy2 = PICH(0) - PICRes.Height + picSmall.Top + PICH(1) - sorcYRes
      sy1 = sy2 - PICH(1)
      picS.Left = sx1 * (picL.Width / PICW(0))
      picS.Top = sy1 * (picL.Height / PICH(0))
   End Select
  
   mnuFile(0).Enabled = True
   
   aRun = True
   cmdDisplay.Enabled = True
End Sub


Private Sub optAdd_Click(Index As Integer)
   AlphaFactor = Index
End Sub

Private Sub optSelect_Click(Index As Integer)
   Select Case Index
   Case 0   ' Subtract
      aSelect(0) = True
      aSelect(1) = False
      Frame2.Visible = False
      Frame1.Left = 390
      Frame1.Visible = True   ' Left = 390
      If aPic0EqPic1 Then
         picSmall.Visible = False
         ReDim ARRRes(1 To PICW(0), 1 To PICH(0))
      Else
         picSmall.Visible = True
         picSmall.Picture = LoadPicture(InFileSpec$(1))
         ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
      End If
      
   Case 1   ' Add
      aSelect(0) = False
      aSelect(1) = True
      Frame1.Visible = False
      Frame2.Left = 390
      Frame2.Visible = True   ' Left = 390
      picSmall.Visible = True
      picSmall.Picture = LoadPicture(InFileSpec$(1))
      ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
   End Select
End Sub

Private Sub optMode_Click(Index As Integer)
   TheMode = Index
End Sub

Private Sub chkInvert_Click()
   InvertYN = 1 - InvertYN
End Sub


'#### SCROLLING ####

Private Sub SetScrollBars()
Dim k As Long
' HS Max/Min
' VS Max/Min
   aSBarsActive = False
   For k = 0 To 1
      If PICW(k) > PICWOrg Then
         HSMin(k) = 0
         HSMax(k) = (PICW(k) - PICWOrg) + 1
         sorcX(k) = HSMin(k)
      Else
         sorcX(k) = HSMin(k)
      End If
      
      If PICH(k) > PICHOrg Then
         VSMax(k) = 0
         VSMin(k) = (PICH(k) - PICHOrg) + 1
         sorcY(k) = VSMin(k)
      Else
         VSMax(k) = 0
         VSMin(k) = (PICH(k) - PICHOrg) + 1
         sorcY(k) = VSMin(k)
      End If
   Next k
   
   SetResScrollBars
   aSBarsActive = True
End Sub

Private Sub SetResScrollBars()
Dim k As Long
   HSResMin = 0
   HSResMax = PICW(0) - PICRes.Width + 1
   sorcXRes = HSResMin

   VSResMax = 0
   VSResMin = (PICH(0) - PICRes.Height) + 1
   sorcYRes = VSResMin
   Borders
End Sub

Private Sub HSAlpha_Change()
   Call HSAlpha_Scroll
End Sub

Private Sub HSAlpha_Scroll()
   AF = HSAlpha.Value
   LabAF = Str$(AF)
   AF = 1.28 * AF
   zAlpha = AF / 128
   If aSBarsActive And aRun Then
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARR0(), sorcXRes, sorcYRes, -1
      ASM_DigSub Me
      DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   End If
End Sub

Private Sub HSWeighting_Change()
   Call HSWeighting_Scroll
End Sub

Private Sub HSWeighting_Scroll()
   Weighting = HSWeighting.Value
   LabWeighting = Weighting
   If aSBarsActive And aRun Then
      picSmall.Visible = True
      picSmall.Picture = LoadPicture(InFileSpec$(1))
      ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
      ASM_DigSub Me
      DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   End If
End Sub

Private Sub HSGreyLevel_Change()
   Call HSGreyLevel_Scroll
End Sub

Private Sub HSGreyLevel_Scroll()
   GreyLevel = HSGreyLevel.Value
   LabGreyLevel = GreyLevel
   If aSBarsActive And aRun Then
      picSmall.Visible = True
      picSmall.Picture = LoadPicture(InFileSpec$(1))
      ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
      ASM_DigSub Me
      DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   End If
End Sub

Private Sub PIC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = True
   PIC(Index).MousePointer = vbCustom
   PIC(Index).MouseIcon = LoadResPicture("SMALLCLOSEDHAND", vbResCursor)
End Sub

Private Sub PIC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If aMouseDown Then
      HandScroller x, y
   End If
End Sub

Private Sub PIC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = False
   PIC(Index).MousePointer = vbCustom
   PIC(Index).MouseIcon = LoadResPicture("SMALLOPENHAND", vbResCursor)
End Sub

Private Sub PICRes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = True
   PICRes.MousePointer = vbCustom
   PICRes.MouseIcon = LoadResPicture("CLOSEDHAND", vbResCursor)
End Sub

Private Sub PICRes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sx1 As Long
Dim sy1 As Long, sy2 As Long
   If aMouseDown Then
      HandScrollerRes x, y
      If aSBarsActive And aRun Then
         picSmall.Visible = True
         picSmall.Picture = LoadPicture(InFileSpec$(1))
         ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
         ARRRes() = ARR1()
         ASM_DigSub Me
         DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
         ' Small view
         sx1 = picSmall.Left + sorcXRes
         sy2 = PICH(0) - PICRes.Height + picSmall.Top + PICH(1) - sorcYRes
         sy1 = sy2 - PICH(1)
         picS.Left = sx1 * (picL.Width / PICW(0))
         picS.Top = sy1 * (picL.Height / PICH(0))
      End If
   End If
End Sub

Private Sub PICRes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PICRes.MousePointer = vbCustom
   PICRes.MouseIcon = LoadResPicture("OPENHAND", vbResCursor)
   aMouseDown = False
End Sub

Private Sub HandScrollerRes(x As Single, y As Single)
'Public HSResMax As Long, HSResMin As Long
'Public VSResMax As Long, VSResMin As Long
Dim sx As Single, sy As Single
   If aSBarsActive And aRun Then
      
      If PICW(0) > PICRes.Width Or PICH(0) > PICRes.Height Then
         sx = x / PICRes.Width * HSResMax
         If sx <= HSResMax Then
         If sx >= HSResMin Then
            sorcXRes = sx
         End If
         End If
         sy = y / PICRes.Height * VSResMin
         If sy <= VSResMin Then
         If sy >= VSResMax Then
            sorcYRes = VSResMin - sy
         End If
         End If
         If aPic0EqPic1 Then
            DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), sorcXRes, sorcYRes, -1
         Else
            DisplayArray PICRes, PICRes.Width, PICRes.Height, ARR0(), sorcXRes, sorcYRes, -1
         End If
      End If
   End If
End Sub

Private Sub HandScroller(x As Single, y As Single)
' Public HSMax As Long, HSMin As Long
' Public VSMax As Long, VSMin As Long
' PIC(0), PIC(1)
Dim sx As Single, sy As Single

   If aSBarsActive Then
      sx = x / PIC(0).Width * HSMax(0)
      If sx <= HSMax(0) Then
      If sx >= HSMin(0) Then
         sorcX(0) = sx
      End If
      End If
      sy = y / PIC(0).Height * VSMin(0)
      If sy <= VSMin(0) Then
      If sy >= VSMax(0) Then
         sorcY(0) = VSMin(0) - sy
      End If
      End If
      DisplayArray PIC(0), PIC(0).Width, PIC(0).Height, ARR0(), sorcX(0), sorcY(0), -1
      If PICW(1) > PICWOrg Then
         sx = x / PIC(1).Width * HSMax(1)
         If sx <= HSMax(1) Then
         If sx >= HSMin(1) Then
            sorcX(1) = sx
         End If
         End If
      Else
         sorcX(1) = 0
      End If

      If PICH(1) > PICHOrg Then
         sy = y / PIC(1).Height * VSMin(0)
         If sy <= VSMin(1) Then
         If sy >= VSMax(1) Then
            sorcY(1) = VSMin(1) - sy
         End If
         End If
      Else
         sorcY(1) = VSMin(1)
      End If
      DisplayArray PIC(1), PIC(1).Width, PIC(1).Height, ARR1(), sorcX(1), sorcY(1), -1
   End If
End Sub

'#### END SCROLLING ####


Private Sub mnuHelp_Click()
   MsgBox "Digital Image Subtraction or Addition" & vbCr _
        & "by  Robert Rayment 2005" & vbCr & vbCr _
        & "1. Needs screen res >= 1024 x 768." & vbCr _
        & "2. Pic 1 Width & Height must be  < =  Pic 0." & vbCr _
        & "3. Smaller picture needs to be inside larger picture" & vbCr _
        & "    for correct result when hand-scrolling." & vbCr _
        & "4. RunB: Run BASIC routines" & vbCr _
        & "5. RunA: Run ASM routines" & vbCr & vbCr _
        & "Minus:  GreyLevel +/- (RGB0 - RGB1) * Weighting" & vbCr _
        & "Xor:     GreyLevel -/+ (RGB0 XOR RGB1) * Weighting " & vbCr _
        & "Invert swaps sign" & vbCr _
        & "Add alpha:    varies transparency of pic 1" & vbCr _
        & "Alpha edges: varies edge transparency of pic 1" _
        , vbInformation, "Info"
End Sub

Private Sub PositionControls()
'Public Const PICWOrg As Long = 256
'Public Const PICHOrg As Long = 256
'Public Const PICResWOrg As Long = 620
'Public Const PICResHOrg As Long = 500
'Public Const FormWOrg As Long = 14200
'Public Const FormHOrg As Long = 9975

Dim k As Long
   GetExtras Me.BorderStyle
   ' IN:  BStyle = Me.BorderStyle
   ' OUT: Public ExtraBorder, ExtraHeight

   For k = 0 To 1
      With PIC(k)
         .Width = PICWOrg
         .Height = PICHOrg
      End With
   Next k
   PIC(1).Left = PIC(0).Left
   With PICRes
      .Width = PICResWOrg
      .Height = PICResHOrg
   End With
   
   cmdLoad(1).Left = cmdLoad(0).Left
   optMode(1).Top = optMode(0).Top
   chkInvert.Top = optMode(0).Top
   cmdRun(0).Top = optMode(0).Top
   cmdRun(1).Top = optMode(0).Top
   
   With HSWeighting
      .Min = 1
      .Max = 32
      .TabStop = False
      .Value = .Min
      Weighting = .Min
      LabWeighting = Weighting
   End With
   With HSGreyLevel
      .Min = 0
      .Max = 255
      .TabStop = False
      .Value = 128
      GreyLevel = .Value
      LabGreyLevel = GreyLevel
   End With
   PICRes.Top = 95
   Borders
End Sub

Private Sub Form_Resize()
'   PICResWDef = 620
'   PICResHDef = 500
'   PICResWOrg = PICResWDef
'   PICResHOrg = PICResHDef
'Public Const FormWOrg As Long = 14200
'Public Const FormHOrg As Long = 9975
Dim sx1 As Long
Dim sy1 As Long, sy2 As Long
Dim k As Long
   If WindowState = vbMinimized Then Exit Sub
   
   If Me.Width <= FormWOrg Or Me.Height <= FormHOrg Then
      Me.Width = FormWOrg
      Me.Height = FormHOrg + ExtraHeight
      PICResWOrg = PICResWDef
      PICResHOrg = PICResHDef
      PICRes.Width = PICResWOrg
      PICRes.Height = PICResHOrg
   ElseIf Me.Width > FormWOrg Or Me.Height > FormHOrg Then
      PICRes.Width = Me.Width / STX - PICRes.Left - 30
      PICRes.Height = Me.Height / STY - PICRes.Top - 80 '100
      PICResWOrg = PICRes.Width
      PICResHOrg = PICRes.Height
   End If
   If aRun Then
      If PICW(0) <= PICRes.Width Then
         PICRes.Width = PICW(0)
      Else
         If PICW(0) <= PICResWOrg Then
            PICRes.Width = PICW(0)
         Else
            PICRes.Width = PICResWOrg
         End If
      End If
      If PICH(0) <= PICRes.Height Then
         PICRes.Height = PICH(0)
      Else
         If PICH(0) <= PICResHOrg Then
            PICRes.Height = PICH(0)
         Else
            PICRes.Height = PICResHOrg
         End If
      End If
      SetScrollBars     ' Also does SetResScrollBars
      ' Small view
      sx1 = picSmall.Left + sorcXRes
      sy2 = PICH(0) - PICRes.Height + picSmall.Top + PICH(1) - sorcYRes
      sy1 = sy2 - PICH(1)
      picS.Left = sx1 * (picL.Width / PICW(0))
      picS.Top = sy1 * (picL.Height / PICH(0))
      
      If aSelect(0) Then
         cmdRun_Click 0
      Else
         cmdRun2_Click 0
      End If
   End If
   Borders
End Sub

Public Sub DisplayArray(p As PictureBox, DW As Long, DH As Long, ARR() As Long, _
   Optional ByVal sorcX As Long = 0, Optional ByVal sorcY As Long = 0, Optional ByVal Mul As Long = 1)
' DisplayArray PIC, DW, DH, ARR(),sorcX, sorcY, Mul
' Public ARR0(), ARR1(), ARRres()
' DW & DH = Display W & H : -
' sorcX & sorcY Scrollbar values

' Image in ARR()
Dim UX As Long, UY As Long
Dim BS As BITMAPINFO
UX = UBound(ARR(), 1)
UY = UBound(ARR(), 2)
   With BS.bmi
      .biSize = 40
      .biwidth = UX
      .biheight = UY * Mul ' if -1 required
      .biPlanes = 1
      .biBitCount = 32    ' Sets up 32-bit colors
      .biCompression = 0
      .biSizeImage = 0 'UX * UY * 4
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   p.Picture = LoadPicture
      
   SetStretchBltMode p.hdc, 4  ' NB Of Dest picbox
   If StretchDIBits(p.hdc, 0, 0, DW, DH, sorcX, sorcY, _
      DW - 1, DH - 1, ARR(1, 1), BS, DIB_PAL_COLORS, vbSrcCopy) = 0 Then
         MsgBox "StretchDIBits Error", vbCritical, " "
   End If
   p.Refresh
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim ix As Long, iy As Long, ixx As Long, iyy As Long
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
Dim CulR As Long, CulG As Long, CulB As Long
   Select Case Index
   Case 0   ' Save Result As 24bpp BMP
      ReDim ARRT(1 To PICW(0), 1 To PICH(0))
      Filt$ = "BMP(*.bmp)|*.bmp"
      FileSpec$ = ""
      Title$ = "Save Result 24bpp BMP"
      InDir$ = CurrPath$ 'Pathspec$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) > 0 Then
         With PIC_Save
            .Width = PICW(0)
            .Height = PICH(0)
         End With
         ARRT() = ARR0()
         ' Copy ARRRes() rectangle to ARR0()
         GetPublicCoords Me
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
         DisplayArray PIC_Save, PICW(0), PICH(0), ARRT(), 0, 0, -1 'Mul
         SavePicture PIC_Save.Image, FileSpec$
         PIC_Save.Picture = LoadPicture
         With PIC_Save
            .Width = 6
            .Height = 6
         End With
         Erase ARRT()
      End If
   Case 1   ' Break
   Case 2   ' Exit
      Form_QueryUnload 1, 0
   End Select
End Sub


Private Sub picSmall_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = True
   picSmallX = x
   picSmallY = y
End Sub

Private Sub picSmall_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sx1 As Long
Dim sy1 As Long, sy2 As Long
Dim pLeft As Long, pTop As Long
   If aMouseDown Then
'  This code shows results as picSmall is move but has trails.
'      pLeft = picSmall.Left + (x - picSmallX)
'      If pLeft < -3 * picSmall.Width / 4 Then pLeft = -3 * picSmall.Width / 4
'      If pLeft > PICRes.Width - picSmall.Width / 4 Then
'         pLeft = PICRes.Width - picSmall.Width / 4
'      End If
'      picSmall.Left = pLeft
'
'      pTop = picSmall.Top + (y - picSmallY)
'      If pTop < -3 * picSmall.Height / 4 Then pTop = -3 * picSmall.Height / 4
'      If pTop > PICRes.Height - picSmall.Height / 4 Then
'         pTop = PICRes.Height - picSmall.Height / 4
'      End If
'      picSmall.Top = pTop
''''
      
'  This has no trails but doesn't show result until button released
      ReleaseCapture
      SendMessage picSmall.hWnd, WM_NCLBUTTONDOWN, 2, 0&
      aMouseDown = False
      
      pLeft = picSmall.Left
      If pLeft < -3 * picSmall.Width / 4 Then
         pLeft = -3 * picSmall.Width / 4
         picSmall.Left = pLeft
      ElseIf pLeft > PICRes.Width - picSmall.Width / 4 Then
         pLeft = PICRes.Width - picSmall.Width / 4
         picSmall.Left = pLeft
      End If
      pTop = picSmall.Top
      If pTop < -3 * picSmall.Height / 4 Then
         pTop = -3 * picSmall.Height / 4
         picSmall.Top = pTop
      ElseIf pTop > PICRes.Height - picSmall.Height / 4 Then
         pTop = PICRes.Height - picSmall.Height / 4
         picSmall.Top = pTop
      End If
''''

      If aSBarsActive And aRun Then
         picSmall.Visible = True
         picSmall.Picture = LoadPicture(InFileSpec$(1))
         ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
         ARRRes() = ARR1()
         ASM_DigSub Me
         DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
      End If
      ' Small view
      sx1 = picSmall.Left + sorcXRes
      sy2 = PICH(0) - PICRes.Height + picSmall.Top + PICH(1) - sorcYRes
      sy1 = sy2 - PICH(1)
      picS.Left = sx1 * (picL.Width / PICW(0))
      picS.Top = sy1 * (picL.Height / PICH(0))
      
   End If
End Sub

Private Sub picSmall_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = False
End Sub

Private Sub RunBASIC()
' Pic 0 < Pic 1 size
' Formulae:
' Mode 0:  GreyLevel +/- (RGB0 - RGB1) * Weighting
' Mode 1:  GreyLevel -/+ (RGB0 XOR RGB1) * Weighting
'          Invert swaps sign
Dim ix As Long, iy As Long, ixx As Long, iyy As Long
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
Dim CulR As Long, CulG As Long, CulB As Long
Dim DiffMul As Long
Dim WD As Long
   'On Error Resume Next
   GetPublicCoords Me
   
   Select Case TheMode
   Case 0
      DiffMul = 1
      If InvertYN = 1 Then DiffMul = -1
      WD = Weighting * DiffMul
      
      iyy = 1
      
      For iy = iy1 To iy2
         ixx = 1
         For ix = ix1 To ix2
            CulR = ARR0(ix, iy)
            CulG = ARR1(ixx, iyy)
            If CulR <> CulG Then
               R0 = (CulR And &HFF&)
               G0 = (CulR And &HFF00&) / &H100&
               B0 = (CulR And &HFF0000) / &H10000
               R1 = (CulG And &HFF&)
               G1 = (CulG And &HFF00&) / &H100&
               B1 = (CulG And &HFF0000) / &H10000
               
               CulR = (GreyLevel + (1& * R0 - R1) * WD)
               CulG = (GreyLevel + (1& * G0 - G1) * WD)
               CulB = (GreyLevel + (1& * B0 - B1) * WD)
              
               If CulR < 0 Then CulR = 0
               If CulG < 0 Then CulG = 0
               If CulB < 0 Then CulB = 0
               ARRRes(ixx, iyy) = RGB(CulR, CulG, CulB)
            Else
               ARRRes(ixx, iyy) = RGB(GreyLevel, GreyLevel, GreyLevel)
            End If
            
            ixx = ixx + 1
            If ixx > UX Then Exit For
         Next ix
         iyy = iyy + 1
         If iyy > UY Then Exit For
      Next iy
   
   Case 1   ' XOR: 0 0 = 0, 1 1 = 0, else 1
      DiffMul = -1
      If InvertYN = 0 Then DiffMul = 1
      WD = Weighting * DiffMul
      
      iyy = 1
            
      For iy = iy1 To iy2
         ixx = 1
         For ix = ix1 To ix2
            CulR = ARR0(ix, iy)
            CulG = ARR1(ixx, iyy)
            If CulR <> CulG Then
               R0 = (CulR And &HFF&)
               G0 = (CulR And &HFF00&) / &H100&
               B0 = (CulR And &HFF0000) / &H10000
               R1 = (CulG And &HFF&)
               G1 = (CulG And &HFF00&) / &H100&
               B1 = (CulG And &HFF0000) / &H10000
               
               CulR = (GreyLevel - (R0 Xor R1) * WD)
               CulG = (GreyLevel - (G0 Xor G1) * WD)
               CulB = (GreyLevel - (B0 Xor B1) * WD)
            
               If CulR < 0 Then CulR = 0
               If CulG < 0 Then CulG = 0
               If CulB < 0 Then CulB = 0
               ARRRes(ixx, iyy) = RGB(CulR, CulG, CulB)
            Else
               ARRRes(ixx, iyy) = RGB(GreyLevel, GreyLevel, GreyLevel)
            End If
            ixx = ixx + 1
            If ixx > UX Then Exit For
         Next ix
         iyy = iyy + 1
         If iyy > UY Then Exit For
      Next iy
   End Select
End Sub


'#### Digital Addition ####

Private Sub cmdRun2_Click(Index As Integer)
'Public Const PICResWOrg As Long = 620
'Public Const PICResHOrg As Long = 480
Dim sx1 As Long
Dim sy1 As Long, sy2 As Long
   ' To Set PICRes size as image
   If PICW(0) <= PICRes.Width Then
      PICRes.Width = PICW(0)
   Else
      If PICW(0) <= PICResWOrg Then
         PICRes.Width = PICW(0)
      Else
         PICRes.Width = PICResWOrg
      End If
   End If
   If PICH(0) <= PICRes.Height Then
      PICRes.Height = PICH(0)
   Else
      If PICH(0) <= PICResHOrg Then
         PICRes.Height = PICH(0)
      Else
         PICRes.Height = PICResHOrg
      End If
   End If
   PICRes.Picture = LoadPicture  ' To Flash
   PICRes.Refresh
   
   picSmall.Visible = True
   picSmall.Picture = LoadPicture(InFileSpec$(1))
   ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
   
   If aRun = False Then
      SetScrollBars     ' Also does SetResScrollBars
   End If
   
   DisplayArray PICRes, PICRes.Width, PICRes.Height, ARR0(), sorcXRes, sorcYRes, -1
   
   Select Case Index
   Case 0    ' Run Basic
      If AlphaFactor = 0 Then
         AlphaImage
      Else
         AlphaEdges
      End If
   Case 1   ' Run ASM
      ARRRes() = ARR1()
      ASM_DigSub Me
   Case 2   ' 0,0
      picSmall.Left = 0
      picSmall.Top = 0
      If aSBarsActive And aRun Then
         picSmall.Visible = True
         'picSmall.Picture = LoadPicture(InFileSpec$(1))
         ReDim ARRRes(1 To PICW(1), 1 To PICH(1))
         ARRRes() = ARR1()
         ASM_DigSub Me
      End If
      ' Small view
      sx1 = picSmall.Left + sorcXRes
      sy2 = PICH(0) - PICRes.Height + picSmall.Top + PICH(1) - sorcYRes
      sy1 = sy2 - PICH(1)
      picS.Left = sx1 * (picL.Width / PICW(0))
      picS.Top = sy1 * (picL.Height / PICH(0))
   End Select
   
   DisplayArray picSmall, PICW(1), PICH(1), ARRRes(), 0, 0, -1
   
   mnuFile(0).Enabled = True
   
   aRun = True
   cmdDisplay.Enabled = True
End Sub

Private Sub AlphaImage()
Dim ix As Long, iy As Long, ixx As Long, iyy As Long
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
Dim CulR As Long, CulG As Long, CulB As Long
Dim CulR2 As Long, CulG2 As Long, CulB2 As Long
   'On Error Resume Next
   GetPublicCoords Me
   
   iyy = 1
      
   For iy = iy1 To iy2 + 1
      ixx = 1
      For ix = ix1 To ix2
         CulR = ARR0(ix, iy)
         CulG = ARR1(ixx, iyy)
        
         R0 = (CulR And &HFF&)
         G0 = (CulR And &HFF00&) / &H100&
         B0 = (CulR And &HFF0000) / &H10000
        
         R1 = (CulG And &HFF&)
         G1 = (CulG And &HFF00&) / &H100&
         B1 = (CulG And &HFF0000) / &H10000
         'zAlpha = AF / 128
         CulR = zAlpha * (1& * R1 - R0) + R0
         CulG = zAlpha * (1& * G1 - G0) + G0
         CulB = zAlpha * (1& * B1 - B0) + B0
         
         ' ASM TEST
         'CulR2 = AF * (1& * R1 - R0)
         'CulG2 = AF * (1& * G1 - G0)
         'CulB2 = AF * (1& * B1 - B0)
         'CulR2 = CulR2 \ 128 + R0
         'CulG2 = CulG2 \ 128 + G0
         'CulB2 = CulB2 \ 128 + B0
         
         If CulR < 0 Then CulR = 0
         If CulG < 0 Then CulG = 0
         If CulB < 0 Then CulB = 0
            
         ARRRes(ixx, iyy) = RGB(CulR, CulG, CulB)
        ixx = ixx + 1
        If ixx > UX Then Exit For
      Next ix
      iyy = iyy + 1
      If iyy > UY Then Exit For
   Next iy
End Sub

Private Sub AlphaEdges()
Dim ix As Long, iy As Long, ixx As Long, iyy As Long
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
Dim CulR As Long, CulG As Long, CulB As Long
Dim iup As Long, jup As Long
Dim StepAlpha As Long
Dim ixs As Long, iys As Long
Dim Alpha As Long

' To match ASM
Dim ix2p1 As Long    ' ix2 + 1
Dim iy2p1 As Long    ' iy2 + 1
Dim PHp1 As Long     ' PICH(1) + 1
Dim PWp1 As Long     ' PICW(1) +1

   'On Error Resume Next
   GetPublicCoords Me
   
   ARRRes() = ARR1()
   iup = (PICW(1)) / 2 + 1
   jup = (PICH(1)) / 2 + 1
   If AF = 0 Then
      StepAlpha = 128
   Else
      StepAlpha = 128 \ AF
   End If
   ix2p1 = ix2 + 1
   iy2p1 = iy2 + 1
   PHp1 = PICH(1) + 1
   PWp1 = PICW(1) + 1
   
   ixs = 1
   iyy = 1
      
   Alpha = 0
   
   For iy = iy1 To iy2
      ixx = ixs
      For ix = ix1 To ix2
         ' Top
         CulR = ARR0(ix, iy)
         CulG = ARR1(ixx, iyy)
         GetFeatherColor Alpha, CulR, CulG, CulB
         ARRRes(ixx, iyy) = RGB(CulR, CulG, CulB)
         ' Bottom
         CulR = ARR0(ix, iy2p1 - iyy)
         CulG = ARR1(ixx, PHp1 - iyy)
         GetFeatherColor Alpha, CulR, CulG, CulB
         ARRRes(ixx, PHp1 - iyy) = RGB(CulR, CulG, CulB)
         ixx = ixx + 1
         If ixx > UX Then Exit For
      Next ix
      ix1 = ix1 + 1
      ix2 = ix2 - 1
      
      If ix2 < ix1 Then Exit For
      
      ixs = ixs + 1
      iyy = iyy + 1
      If iyy > jup Then Exit For
      Alpha = Alpha + StepAlpha
      If Alpha > 128 Then Alpha = 128
   Next iy

   ' Left & Right
   GetPublicCoords Me
   
   iys = 1
   ixx = 1
   
   Alpha = 0
   
   For ix = ix1 To ix2
      iyy = iys
      For iy = iy1 To iy2
         ' Left
         CulR = ARR0(ix, iy)
         CulG = ARR1(ixx, iyy)
         GetFeatherColor Alpha, CulR, CulG, CulB
         ARRRes(ixx, iyy) = RGB(CulR, CulG, CulB)
         ' Right
         CulR = ARR0(ix2p1 - ixx, iy)
         CulG = ARR1(PWp1 - ixx, iyy)
         GetFeatherColor Alpha, CulR, CulG, CulB
         ARRRes(PWp1 - ixx, iyy) = RGB(CulR, CulG, CulB)
         iyy = iyy + 1
         If iyy > UY Then Exit For
      Next iy
      iy1 = iy1 + 1
      iy2 = iy2 - 1
      
      If iy2 < iy1 Then Exit For
      
      iys = iys + 1
      ixx = ixx + 1
      If ixx > iup Then Exit For
      Alpha = Alpha + StepAlpha
      If Alpha > 128 Then Alpha = 128
   Next ix
End Sub

Private Sub GetFeatherColor(ALP As Long, CulR As Long, CulG As Long, CulB As Long)
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
   R0 = (CulR And &HFF&)
   G0 = (CulR And &HFF00&) / &H100&
   B0 = (CulR And &HFF0000) / &H10000
   R1 = (CulG And &HFF&)
   G1 = (CulG And &HFF00&) / &H100&
   B1 = (CulG And &HFF0000) / &H10000
   CulR = (128 - ALP) * R0 + ALP * R1
   CulG = (128 - ALP) * G0 + ALP * G1
   CulB = (128 - ALP) * B0 + ALP * B1
   CulR = CulR \ 128
   CulG = CulG \ 128
   CulB = CulB \ 128
   If CulR < 0 Then CulR = 0
   If CulG < 0 Then CulG = 0
   If CulB < 0 Then CulB = 0
End Sub

Private Sub Borders()
Dim k As Long
   Cls
   Line (PICRes.Left - 1, PICRes.Top - 1)-(PICRes.Left + PICRes.Width, PICRes.Top + PICRes.Height), 0, B
   For k = 0 To 1
   Line (PIC(k).Left - 1, PIC(k).Top - 1)-(PIC(k).Left + PIC(k).Width, PIC(k).Top + PIC(k).Height), 0, B
   Next k
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form
Dim res As Long
   If UnloadMode = 0 Then    'Close on Form1 pressed
      res = MsgBox("", vbQuestion + vbYesNo + vbSystemModal, "Quit ?")
      If res = vbNo Then
         Cancel = True
      Else
         Cancel = False
         Screen.MousePointer = vbDefault
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      End If
   End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   Unload Me
'   End
'End Sub


