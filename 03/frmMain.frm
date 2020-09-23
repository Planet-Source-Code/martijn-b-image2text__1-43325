VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image2Text"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMapping 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mapping"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   1695
      Begin VB.OptionButton optRelative 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Relative"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAbsolute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Absolute"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Timer tmrHide 
      Left            =   120
      Top             =   2520
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.bmp"
      DialogTitle     =   "Select a Picture"
      Filter          =   "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
   End
   Begin VB.CommandButton cmdColumns 
      Caption         =   "Set &Width"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangeFont 
      Caption         =   "Change &Font"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoadPicture 
      Caption         =   "&Load Picture"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCopyAscii 
      Caption         =   "&Copy Ascii"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picResized 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3240
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1800
   End
   Begin VB.PictureBox picCharScan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2040
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowASCII 
      Caption         =   "Show &Ascii"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private lCharMinV As Single
Private lCharMaxV As Single
Private arrCharVal(0& To 1&, 0& To 255&) As Single
Private arrCharMap(0& To 255&) As Long
Private strAscii As String
Private lOutputColumns As Long
Private Enum MapMethode
  AbsoluteMapping = 0&
  RelativeMapping = 1&
End Enum
Private lCharMappingMethode As MapMethode


Private Sub cmdChangeFont_Click()

  With dlgMain

    .Flags = cdlCFBoth
    .FontName = picCharScan.FontName
    .FontSize = picCharScan.FontSize
    .FontBold = picCharScan.FontBold
    .FontItalic = picCharScan.FontItalic
    .FontUnderline = picCharScan.FontUnderline
    .ShowFont

    If .FontName <> vbNullString Then

      '***  set the font
      subSetFont picCharScan, .FontName, .FontSize
      '***  get the values for this font
      subGetCharValues picCharScan
      '***  map the chars to grey values
      subMapChars
      '***  set lOutputColumns to maximum for this screen
      lOutputColumns = Screen.Width \ picCharScan.Width \ Screen.TwipsPerPixelX

    End If

  End With

End Sub


Private Sub cmdColumns_Click()

  lOutputColumns = CLng("0" & InputBox("Enter n the line length for the ASCII output", , lOutputColumns))

  If lOutputColumns = 0& Then lOutputColumns = 80&

End Sub


Private Sub cmdCopyAscii_Click()

  If strAscii <> vbNullString Then

    Clipboard.Clear
    Clipboard.SetText strAscii
    MsgBox "ASCII picture copied to the clipboard." & vbNewLine & _
       "Use " & FontName & " Size " & FontSize

  End If

End Sub


Private Sub cmdLoadPicture_Click()

  With dlgMain

    .ShowOpen

    If .FileName <> vbNullString Then

      picOriginal.AutoSize = True
      picOriginal.Picture = LoadPicture(.FileName)

    End If

  End With

End Sub


Private Sub cmdShowASCII_Click()

  '***  scale the picture to font aspect ratio

  With picResized

    .Width = lOutputColumns
    '***  correct for column resize and for character aspect ratio
    .Height = picOriginal.Height / (picOriginal.Width / lOutputColumns) / (picCharScan.Height / picCharScan.Width)
    .PaintPicture picOriginal.Image, 0, 0, .Width, .Height, _
       0, 0, picOriginal.Width, picOriginal.Height, vbSrcCopy

  End With

  '***  now get the text for picOriginal
  strAscii = fcnGetAscii(picResized)
  '***  print the ascii
  subSetFont Me, picCharScan.FontName, picCharScan.FontSize
  Cls
  Print strAscii

End Sub


Private Function fcnGetAscii(ctl As PictureBox) As String

  '***  returns the ascii picture
  '***  arrCharMap() needs to contain the ascii values for each grey value
  Dim lX As Long, lY As Long
  Dim strLine() As String
  Dim lColor As Long
  ctl.ScaleMode = vbPixels
  ReDim strLine(0& To ctl.ScaleHeight - 1&) As String

  With ctl

    For lY = 0& To .ScaleHeight - 1&

      strLine(lY) = String$(.ScaleWidth, " ")

      For lX = 0& To .ScaleWidth - 1&

        '***  read the color from the picture
        lColor = GetPixel(.hDC, lX, lY)
        '***  get the character for this color
        Mid$(strLine(lY), lX + 1&, 1&) = Chr$(arrCharMap(fcnGetGrey(lColor)))

      Next

    Next

  End With

  '***  return the ascii picture
  fcnGetAscii = Join(strLine, vbNewLine)

End Function


Private Sub subMapChars()

  '***  maps grey values to characters
  '***  arrCharMap() is filled
  '***  arrCharMap(127) stores the ascii value for grey value 127
  '***  color = RGB(127,127,127) = 8355711
  Dim lCount As Long

  Select Case lCharMappingMethode

    Case 0&
      '***  using ABSOLUTE matching

      For lCount = 0& To 255&

        arrCharMap(lCount) = fcnGetClosestChar(lCount)

      Next

    Case 1&
      '***  using relative matching
      '***  this will result in better character variation, but less contrast.
      '***  store the _original_ position = ascii value

      For lCount = 0& To 30&

        arrCharVal(1&, lCount) = 32&

      Next

      For lCount = 31& To 255&

        arrCharVal(1&, lCount) = lCount

      Next

      '***  use simple bubble-sort to sort the characters from dark to light
      subSortCharVal

      For lCount = 0& To 255&

        arrCharMap(lCount) = arrCharVal(1&, lCount)

      Next

  End Select

End Sub


Private Sub subSortCharVal()

  Dim lCount As Long
  Dim lCompare As Long
  Dim lSwap As Variant

  For lCount = 0& To 255&

    For lCompare = lCount To 255&

      If arrCharVal(0&, lCompare) < arrCharVal(0&, lCount) Then

        '***  swap the values
        lSwap = arrCharVal(0&, lCompare)
        arrCharVal(0&, lCompare) = arrCharVal(0&, lCount)
        arrCharVal(0&, lCount) = lSwap
        lSwap = arrCharVal(1&, lCompare)
        arrCharVal(1&, lCompare) = arrCharVal(1&, lCount)
        arrCharVal(1&, lCount) = lSwap

      End If

    Next

  Next

End Sub


Private Function fcnGetClosestChar(lValue As Long) As Long

  '***  finds the best character (=ascii value) for a grey value
  '***  results depend on the font 'grey density' distribution
  Dim lCount As Long
  Dim sOptimum As Single
  Dim sRange As Single
  Dim sTargetVal As Single
  sOptimum = 255
  sRange = lCharMaxV - lCharMinV
  '***  try to find a character closest to sTargetVal
  '***  sTargetVal is scaled between the darkest and lightest character found
  sTargetVal = lCharMinV + (lValue / 255) * sRange

  For lCount = 0& To 255&

    '***  test all characters

    If Abs(arrCharVal(0&, lCount) - sTargetVal) < sOptimum Then

      '***  store the best value found
      sOptimum = Abs(arrCharVal(0&, lCount) - sTargetVal)
      '***  return the ascii value
      fcnGetClosestChar = lCount

    End If

  Next

End Function


Private Sub subSetFont(ctl As Object, sFont As String, lSize As Long)

  '***  set the font size and name

  With ctl

    .FontName = sFont
    .FontSize = lSize

  End With

End Sub


Private Sub subGetCharValues(ctl As PictureBox)

  '***  stores the grey values for all characters in arrCharVal()
  Dim lCount As Long
  Dim szChars As SizeAPI
  '***  reset the minimum and maximum value
  lCharMaxV = -1&
  lCharMinV = 256&

  With ctl

    .BorderStyle = 0&
    .Appearance = 0&
    '***  get the font dimensions. this used to be
    '***  .Width = .TextWidth("A")
    '***  .Height = .TextHeight("A")
    '***  now I use this for better results with the Terminal Font.
    szChars = GetTextSize(picCharScan.Font, "A")
    .Width = szChars.Width
    .Height = szChars.Height
    .AutoRedraw = True
    .BackColor = vbWhite
    .ForeColor = vbBlack
    '***  this will set the range of characters used
    '***  for standard ascii change:
    '***  For lCount = to 31& to 127&

    For lCount = 0& To 30&

      arrCharVal(0&, lCount) = 99999

    Next

    For lCount = 31& To 255&

      .Cls
      ctl.Print Chr$(lCount);
      arrCharVal(0&, lCount) = fcnGetAvg(ctl)
      '***  store the minimum and maximum value for later use

      If arrCharVal(0&, lCount) > lCharMaxV Then lCharMaxV = arrCharVal(0&, lCount)

      If arrCharVal(0&, lCount) < lCharMinV Then lCharMinV = arrCharVal(0&, lCount)

    Next

  End With

End Sub


Private Function fcnGetAvg(ctl As PictureBox) As Single

  '***  returns the grey value of a black and white picture
  '***  this is used to get the grey value of a single character
  '***  grey value ranges 0 to 255 (white to black)
  Dim lX As Long, lY As Long
  Dim lCT As Long

  With ctl

    .ScaleMode = vbPixels
    '***  count the number of white pixels

    For lX = 0& To .ScaleWidth - 1

      For lY = 0& To .ScaleHeight - 1

        If GetPixel(.hDC, lX, lY) = vbWhite Then

          lCT = lCT + 1&

        End If

      Next

    Next

  End With

  '***  return: ( white pixels / total pixels ) * 255
  fcnGetAvg = (lCT / (lX * lY)) * 255&

End Function


Private Function fcnGetGrey(lCol As Long) As Long

  '***  simply returns the grey value for a color
  '***  grey value ranges 0 to 255 (white to black)
  '***  grey color = RGB(GV, GV, GV)
  Dim lR As Long, lG As Long, lB As Long
  lR = Abs(lCol Mod 256&)
  lG = Abs((lCol Mod 65536) \ 256&)
  lB = Abs(lCol \ 65536)
  '***  return the AVG of R+G+B
  fcnGetGrey = (lR + lG + lB) / 3&

End Function


Private Sub Form_Load()

  '***  set the default OutputSize
  lOutputColumns = 240&
  '***  set the default Mapping
  lCharMappingMethode = AbsoluteMapping
  '***  set the default font
  subSetFont picCharScan, "Lucida Console", 6
  '***  get the values for this font
  subGetCharValues picCharScan
  '***  map the chars to grey values
  subMapChars

End Sub


Private Sub optAbsolute_Click()

  lCharMappingMethode = AbsoluteMapping
  '***  get the values for this font
  subGetCharValues picCharScan
  '***  map the chars to grey values
  subMapChars

End Sub


Private Sub optRelative_Click()

  lCharMappingMethode = RelativeMapping
  '***  get the values for this font
  subGetCharValues picCharScan
  '***  map the chars to grey values
  subMapChars

End Sub


'*** code below is just to hide the picture when moving the mouse over it
Private Sub picOriginal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  '***  hide the picture for a second..

  With tmrHide

    .Interval = 1000
    .Enabled = True
    picOriginal.Visible = False

  End With

End Sub


Private Sub tmrHide_Timer()

  With tmrHide

    .Interval = 0
    .Enabled = False

  End With

  picOriginal.Visible = True

End Sub

