VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Icon Editor"
   ClientHeight    =   5580
   ClientLeft      =   36
   ClientTop       =   1020
   ClientWidth     =   9636
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   7.8
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   6132
   Icon            =   "FRM1ICN4.frx":0000
   Left            =   -12
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   803
   Top             =   516
   Width           =   9732
   Begin VB.Timer TimerRefresh 
      Left            =   2520
      Top             =   3000
   End
   Begin VB.CommandButton CmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox ChkSolidBox 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solid"
      Enabled         =   0   'False
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   4800
      Width           =   855
   End
   Begin VB.CheckBox ChkSolidCirc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solid"
      Enabled         =   0   'False
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.VScrollBar VScrollRed 
      Height          =   3870
      LargeChange     =   1280
      Left            =   3360
      SmallChange     =   128
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar VScrollGreen 
      Height          =   3870
      LargeChange     =   1280
      Left            =   4080
      SmallChange     =   128
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar VScrollBlue 
      Height          =   3840
      LargeChange     =   1280
      Left            =   4800
      SmallChange     =   128
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text2 - Debug!"
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "UNTITLED"
      Top             =   0
      Width           =   4572
   End
   Begin VB.Label LblRedSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label LblGreenSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      ForeColor       =   &H0000BF00&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label LblBlueSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image BoxOn 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":030A
      Top             =   4440
      Width           =   540
   End
   Begin VB.Image BoxOff 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":07C4
      Top             =   4440
      Width           =   540
   End
   Begin VB.Image PenOff 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":0C7E
      Top             =   600
      Width           =   540
   End
   Begin VB.Image PenOn 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":1138
      Top             =   600
      Width           =   540
   End
   Begin VB.Label LblRedVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblGreenVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H0000BF00&
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblBlueVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblColr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   6720
      TabIndex        =   15
      Top             =   5280
      Width           =   84
   End
   Begin VB.Label LblXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   14
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   552
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   112
      Y2              =   160
   End
   Begin VB.Label LblPal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palette"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   15
      Left            =   1920
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   14
      Left            =   1560
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   13
      Left            =   1920
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   12
      Left            =   1560
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   11
      Left            =   1920
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   10
      Left            =   1560
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   9
      Left            =   1920
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   8
      Left            =   1560
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   7
      Left            =   1920
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   1560
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   1920
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   1560
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   1920
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   1560
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   1920
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Pall 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   1560
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Box"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ellipse"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   588
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Line"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fill"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   270
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Draw"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   435
   End
   Begin VB.Image CircOn 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":15F2
      Top             =   3480
      Width           =   540
   End
   Begin VB.Image CircOff 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":1AAC
      Top             =   3480
      Width           =   540
   End
   Begin VB.Image LineOn 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":1F66
      Top             =   2520
      Width           =   540
   End
   Begin VB.Image LineOff 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":2420
      Top             =   2520
      Width           =   540
   End
   Begin VB.Image FillOn 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":28DA
      Top             =   1560
      Width           =   540
   End
   Begin VB.Image FillOff 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   120
      Picture         =   "FRM1ICN4.frx":2D94
      Top             =   1560
      Width           =   540
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu Menu 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu 
         Caption         =   "S&ave as..."
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu Menu 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu ImgMenu 
      Caption         =   "&Options"
      Begin VB.Menu Img 
         Caption         =   "Black isTransparent"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu Help 
         Caption         =   "About Icon Editor..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub BoxOff_Click()
If SelectedColor < -0.5 Then Exit Sub
If CircStartFlag = 1 Then Exit Sub
If LineStartFlag = 1 Then Exit Sub
PenFlag = 0: PenOff.Visible = True: PenOn.Visible = False
FillFlag = 0: FillOff.Visible = True: FillOn.Visible = False
LineFlag = 0: LineOff.Visible = True: LineOn.Visible = False
CircFlag = 0: CircOff.Visible = True: CircOn.Visible = False
ChkSolidCirc.Enabled = False
BoxFlag = 1: BoxOff.Visible = False: BoxOn.Visible = True
ChkSolidBox.Enabled = True
If ChkSolidBox.Value = 1 Then Form1.FillStyle = 0: Form1.FillColor = Pal(SelectedColor)
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = True
End Sub

Private Sub CircOff_Click()
If SelectedColor < -0.5 Then Exit Sub
Rem If CircStartFlag = 1 Then Exit Sub
If LineStartFlag = 1 Then Exit Sub
PenFlag = 0: PenOff.Visible = True: PenOn.Visible = False
FillFlag = 0: FillOff.Visible = True: FillOn.Visible = False
LineFlag = 0: LineOff.Visible = True: LineOn.Visible = False
CircFlag = 1: CircOff.Visible = False: CircOn.Visible = True
ChkSolidCirc.Enabled = True
If ChkSolidCirc.Value = 1 Then Form1.FillStyle = 0: Form1.FillColor = Pal(SelectedColor)
BoxFlag = 0: BoxOff.Visible = True: BoxOn.Visible = False
ChkSolidBox.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = True
Label5.Enabled = False
End Sub

Private Sub CmdRefresh_Click()
Call DrawIcon
End Sub

Private Sub FillOff_Click()
If SelectedColor < -0.5 Then Exit Sub
If CircStartFlag = 1 Then Exit Sub
If LineStartFlag = 1 Then Exit Sub
PenFlag = 0: PenOff.Visible = True: PenOn.Visible = False
FillFlag = 1: FillOff.Visible = False: FillOn.Visible = True
LineFlag = 0: LineOff.Visible = True: LineOn.Visible = False
CircFlag = 0: CircOff.Visible = True: CircOn.Visible = False
ChkSolidCirc.Enabled = False
BoxFlag = 0: BoxOff.Visible = True: BoxOn.Visible = False
ChkSolidBox.Enabled = False
Label1.Enabled = False
Label2.Enabled = True
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
End Sub

Private Sub Form_Load()
If ExitFlag = 1 Then Exit Sub
Rem Visible = True
Width = Int(8376 * Screen.TwipsPerPixelX / 12 + 0.5)
Rem Width = Int(6516 * screen.TwipsPerPixelX / 12 + .5)
Height = Int(5076 * Screen.TwipsPerPixelY / 12 + 0.5)
Rem ScaleWidth = 535
For a% = 0 To 7: Pos(a%) = 2 ^ (7 - a%): Next a%
ScaleWidth = 690
ScaleHeight = 377
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
LblPal.Left = 143
LblPal.Top = 344
PenOff.Left = 10
PenOff.Top = 70
PenOn.Left = 10
PenOn.Top = 70
FillOn.Left = 10
FillOn.Top = 128
FillOff.Left = 10
FillOff.Top = 128
LineOn.Left = 10
LineOn.Top = 186
LineOff.Left = 10
LineOff.Top = 186
CircOn.Left = 10
CircOn.Top = 244
CircOff.Left = 10
CircOff.Top = 244
ChkSolidCirc.Left = 65
ChkSolidCirc.Top = 279
BoxOn.Left = 10
BoxOn.Top = 302
BoxOff.Left = 10
BoxOff.Top = 302
ChkSolidBox.Left = 65
ChkSolidBox.Top = 337
Text1.Left = 90
Text1.Top = 10
Text1.Height = 20
Text1.Width = 560
LblName.Left = 20
LblName.Top = 13
Label1.Left = 70
Label1.Top = 85
Label2.Left = 70
Label2.Top = 143
Label3.Left = 70
Label3.Top = 201
Label4.Left = 70
Label4.Top = 259
Label5.Left = 70
Label5.Top = 317
VScrollRed.Left = 224
VScrollRed.Top = 80
VScrollRed.Height = 248
VScrollGreen.Left = 272
VScrollGreen.Top = 80
VScrollGreen.Height = 248
VScrollBlue.Left = 320
VScrollBlue.Top = 80
VScrollBlue.Height = 248
LblRedSB.Left = 216
LblRedSB.Top = 344
LblRedSB.Height = 18
LblGreenSB.Left = 264
LblGreenSB.Top = 344
LblGreenSB.Height = 18
LblBlueSB.Left = 312
LblBlueSB.Top = 344
LblBlueSB.Height = 18
LblXY.Left = 480
LblXY.Top = 40
LblXY.Width = 145
LblRedVal.Top = 358
LblRedVal.Left = 500
LblRedVal.Height = 20
LblGreenVal.Top = 358
LblGreenVal.Left = 550
LblGreenVal.Height = 20
LblBlueVal.Top = 358
LblBlueVal.Left = 600
LblBlueVal.Height = 20
LblColr.Left = 420
LblColr.Top = 358
CmdRefresh.Top = 40
CmdRefresh.Left = 233
CmdRefresh.Height = 25
CmdRefresh.Width = 97
Rem LblColr.Width = 145:Rem (Autosized)

For a% = 0 To 15
X = 140 + (a% And 1) * 30
Y = 145 + Int(a% / 2) * 25
Pall(a%).Left = X
Pall(a%).Top = Y
Next a%

BoxWidth = 292
BoxHeight = 292
Rem Size of pixel boxes:
SqWidth = Int((BoxWidth - 66) / 32)
SqHeight = Int((BoxHeight - 66) / 32)
BoxX1 = 375
BoxY1 = 60
BoxX2 = BoxX1 + BoxWidth
BoxY2 = BoxY1 + BoxHeight
PreviewX1 = 144
PreviewY1 = 80
PreviewX2 = PreviewX1 + 35
PreviewY2 = PreviewY1 + 35
Call NewFile
Call DrawEditBox
Call DrawPreviewBox
FileChangedFlag = 0: CmdRefresh.Enabled = False: TimerRefresh = False
Form1.Visible = True
Rem 1st 62 bytes of icon file:
Hdr = "0000010001002020100000000000E802"
Hdr = Hdr + "00001600000028000000200000004000"
Hdr = Hdr + "00000100040000000000800200000000"
Hdr = Hdr + "0000000000000000000000000000"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem Check to see if a color is chosen
For n% = 0 To 15
If (X > Pall(n%).Left) And (X < (Pall(n%).Left + Pall(n%).Width)) And (Y > Pall(n%).Top) And (Y < (Pall(n%).Top + Pall(n%).Height)) Then
  SelectedColor = n%
  Call OutlineBox
  Rem Form1.AutoRedraw = True
  Exit Sub
End If
Next n%
If (X < BoxX1 + 3) Or (X > BoxX2 - 3) Or (Y < BoxY1 + 3) Or (Y > BoxY2 - 3) Then Exit Sub
ClickedPixelX = Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32)
ClickedPixelY = Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
If PenFlag = 1 Then
  Line (ClickedPixelX * (SqWidth + 2) + BoxX1 + 3, ClickedPixelY * (SqHeight + 2) + BoxY1 + 3)-((ClickedPixelX + 1) * (SqWidth + 2) - 2 + BoxX1 + 3, (ClickedPixelY + 1) * (SqHeight + 2) - 2 + BoxY1 + 3), Pal(SelectedColor), BF
  PSet (ClickedPixelX + PreviewX1 + 2, ClickedPixelY + PreviewY1 + 2), Pal(SelectedColor)
  PixArray(ClickedPixelX, ClickedPixelY) = SelectedColor
  FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
End If
If FillFlag = 1 Then
  Form1.MousePointer = 11
  ClickedColor% = PixArray(ClickedPixelX, ClickedPixelY)
  PixArray(ClickedPixelX, ClickedPixelY) = 16
SearchPixels:
  PixelsChanged% = 0
  For X = 0 To 31
  For Y = 0 To 31
  If PixArray(X, Y) = 16 Then
    If (X > 0) Then
      If PixArray(X - 1, Y) = ClickedColor% Then PixArray(X - 1, Y) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (Y > 0) Then
      If PixArray(X, Y - 1) = ClickedColor% Then PixArray(X, Y - 1) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (X < 31) Then
      If PixArray(X + 1, Y) = ClickedColor% Then PixArray(X + 1, Y) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (Y < 31) Then
      If PixArray(X, Y + 1) = ClickedColor% Then PixArray(X, Y + 1) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
  End If
  Next Y
  Next X
  If PixelsChanged% > 0 Then GoTo SearchPixels
  For X = 0 To 31
  For Y = 0 To 31
  If PixArray(X, Y) = 16 Then PixArray(X, Y) = SelectedColor
  Next Y
  Next X
  Call DrawIcon
  Form1.MousePointer = 1
  FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
  Exit Sub
End If
If LineFlag = 1 Then
  If LineStartFlag = 0 Then
    LineStartFlag = 1
    LineStartX = ClickedPixelX
    LineStartY = ClickedPixelY
    Exit Sub
  End If
  If LineStartFlag = 1 Then
    LineStartFlag = 0
    LineEndX = ClickedPixelX
    LineEndY = ClickedPixelY
    Line1.Visible = False
    Line (LineStartX + PreviewX1 + 2, LineStartY + PreviewY1 + 2)-(LineEndX + PreviewX1 + 2, LineEndY + PreviewY1 + 2), Pal(SelectedColor)
    PSet (LineEndX + PreviewX1 + 2, LineEndY + PreviewY1 + 2), Pal(SelectedColor)
    Call CopyLittleToBig
    FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
    Exit Sub
  End If
End If
If CircFlag = 1 Then
  If CircStartFlag = 0 Then
    CircStartFlag = 1
    CircStartX = ClickedPixelX
    CircStartY = ClickedPixelY
    Exit Sub
  End If
  If CircStartFlag = 1 Then
    Form1.FillColor = Pal(SelectedColor)
    CircStartFlag = 0
    CircEndX = ClickedPixelX
    CircEndY = ClickedPixelY
    AspectRatio! = 1
    If (CircEndX <> CircStartX) And (CircEndY <> CircStartY) Then AspectRatio! = Abs((CircEndY - CircStartY) / (CircEndX - CircStartX))
    If (CircEndX = CircStartX) Then AspectRatio! = 100
    If (CircEndY = CircStartY) Then AspectRatio! = 0.01
    Radius% = Int(Abs(CircStartX - CircEndX) / 2)
    If Abs(CircStartY - CircEndY) > Abs(CircStartX - CircEndX) Then Radius% = Int(Abs(CircStartY - CircEndY) / 2)
    If ChkSolidCirc.Value = 1 Then Form1.FillStyle = 0
    If ChkSolidCirc.Value = 0 Then Form1.FillStyle = 1
    Circle ((CircStartX + CircEndX) / 2 + PreviewX1 + 2, (CircStartY + CircEndY) / 2 + PreviewY1 + 2), Radius% + 0.5, Pal(SelectedColor), , , AspectRatio!
    Form1.FillStyle = 1
    Call CopyLittleToBig
    Call DrawGrid
    FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
    Exit Sub
  End If
End If
If BoxFlag = 1 Then
  If BoxStartFlag = 0 Then
    BoxStartFlag = 1
    BoxStartX = ClickedPixelX
    BoxStartY = ClickedPixelY
    Exit Sub
  End If
  If BoxStartFlag = 1 Then
    BoxStartFlag = 0
    BoxEndX = ClickedPixelX
    BoxEndY = ClickedPixelY
    If ChkSolidBox.Value = 0 Then Line (BoxStartX + PreviewX1 + 2, BoxStartY + PreviewY1 + 2)-(BoxEndX + PreviewX1 + 2, BoxEndY + PreviewY1 + 2), Pal(SelectedColor), B
    If ChkSolidBox.Value = 1 Then Line (BoxStartX + PreviewX1 + 2, BoxStartY + PreviewY1 + 2)-(BoxEndX + PreviewX1 + 2, BoxEndY + PreviewY1 + 2), Pal(SelectedColor), BF
    Form1.FillStyle = 1
    FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
    Call CopyLittleToBig
    Exit Sub
  End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FileFoundFlag = 0
If (X < BoxX1 + 3) Or (X > BoxX2 - 3) Or (Y < BoxY1 + 3) Or (Y > BoxY2 - 3) Then
  LblXY.Caption = ""
  Rem Text1.Text = FName
  LblColr.Visible = False
  LblBlueVal.Visible = False
  LblGreenVal.Visible = False
  LblRedVal.Visible = False
  Rem CircEndX = Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32)
  Rem CircEndY = Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
  Exit Sub
End If
LblXY.Caption = "X, Y:  " & Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32) & ", " & Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
G1% = PixArray(Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32), Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32))
Rem gwc
b% = Int(Pal(G1%) / 65536) And 255
g% = Int(Pal(G1%) / 256) And 255
r% = Pal(G1%) And 255
LblColr.Caption = "Color# " & G1%
LblBlueVal.Caption = b%
LblGreenVal.Caption = g%
LblRedVal.Caption = r%
LblColr.Visible = True
LblBlueVal.Visible = True
LblGreenVal.Visible = True
LblRedVal.Visible = True
If LineStartFlag = 1 Then
    LineEndX = Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32)
    LineEndY = Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
    Line1.BorderColor = Pal(SelectedColor)
    Line1.X1 = Int((LineStartX + 0.5) * (SqWidth + 2) + BoxX1 + 2)
    Line1.X2 = Int((LineEndX + 0.5) * (SqWidth + 2) + BoxX1 + 2)
    Line1.Y1 = Int((LineStartY + 0.5) * (SqHeight + 2) + BoxY1 + 2)
    Line1.Y2 = Int((LineEndY + 0.5) * (SqHeight + 2) + BoxY1 + 2)
    If Line1.Visible = False Then Line1.Visible = True
    Form1.FillStyle = 1
End If
If CircStartFlag = 1 Then
    CircEndX = Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32)
    CircEndY = Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
    If (PrevCircEndX = CircEndX) And (PrevCircEndY = CircEndY) Then Exit Sub
    Rem Form1.AutoRedraw = True
    PrevCircEndX = CircEndX
    PrevCircEndY = CircEndY
    Rem Form1.AutoRedraw = False
    AspectRatio! = 1
    If (CircEndX <> CircStartX) And (CircEndY <> CircStartY) Then AspectRatio! = Abs((CircEndY - CircStartY) / (CircEndX - CircStartX))
    If (CircEndX = CircStartX) Then AspectRatio! = 100
    If (CircEndY = CircStartY) Then AspectRatio! = 0.01
    Radius% = Int(((Abs(CircStartX - CircEndX) / 2) + 0) * (SqWidth + 2))
    If Abs(CircStartY - CircEndY) > Abs(CircStartX - CircEndX) Then Radius% = (Int(Abs(CircStartY - CircEndY) / 2) + 0.5) * (SqHeight + 2)
    If ChkSolidCirc.Value = 1 Then Form1.FillStyle = 0
    If ChkSolidCirc.Value = 0 Then Form1.FillStyle = 1
    Circle ((((CircStartX + CircEndX) / 2) + 0.5) * (SqWidth + 2) + BoxX1 + 2, (((CircStartY + CircEndY) / 2) + 0.5) * (SqHeight + 2) + BoxY1 + 2), Radius%, Pal(SelectedColor), , , AspectRatio!
    Form1.FillStyle = 1
End If
If BoxStartFlag = 1 Then
    BoxEndX = Int((X - BoxX1 - 2) / (BoxWidth - 4) * 32)
    BoxEndY = Int((Y - BoxY1 - 2) / (BoxHeight - 4) * 32)
    If ChkSolidBox.Value = 1 Then Form1.FillStyle = 0
    If ChkSolidBox.Value = 0 Then Form1.FillStyle = 1
    If ChkSolidBox.Value = 0 Then Line (Int((BoxStartX + 0.5) * (SqWidth + 2) + BoxX1 + 2), Int((BoxStartY + 0.5) * (SqHeight + 2) + BoxY1 + 2))-(Int((BoxEndX + 0.5) * (SqWidth + 2) + BoxX1 + 2), Int((BoxEndY + 0.5) * (SqHeight + 2) + BoxY1 + 2)), Pal(SelectedColor), B
    If ChkSolidBox.Value = 1 Then Line (Int((BoxStartX + 0.5) * (SqWidth + 2) + BoxX1 + 2), Int((BoxStartY + 0.5) * (SqHeight + 2) + BoxY1 + 2))-(Int((BoxEndX + 0.5) * (SqWidth + 2) + BoxX1 + 2), Int((BoxEndY + 0.5) * (SqHeight + 2) + BoxY1 + 2)), Pal(SelectedColor), BF
    Form1.FillStyle = 1
End If
Rem Text6.Text = X
Rem Text7.Text = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ExitFlag = 1 Then Exit Sub
If FileChangedFlag = 1 Then
  a% = MsgBox("Save Current Icon File?", 51, "File NOT Saved")
  If a% = 6 Then
    Rem Yes
    Rem Call SaveFile
    ExitFlag = 1
    Cancel = 1
    Load Form3
    Exit Sub
  End If
  If a% = 7 Then Cancel = 0: Exit Sub: Rem No
  If a% = 2 Then Cancel = 1: Exit Sub: Rem Cancel
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Unload Form3
Unload Form4
End Sub


Private Sub Help_Click(Index As Integer)
Load Form4
End Sub

Private Sub Img_Click(Index As Integer)
If Index = 0 Then
  Rem Black is transparent
  FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True
  Img(0).Checked = -1 - Img(0).Checked
End If
End Sub

Private Sub LineOff_Click()
If SelectedColor < -0.5 Then Exit Sub
If CircStartFlag = 1 Then Exit Sub
Rem If LineStartFlag = 1 Then Exit Sub
PenFlag = 0: PenOff.Visible = True: PenOn.Visible = False
FillFlag = 0: FillOff.Visible = True: FillOn.Visible = False
LineFlag = 1: LineOff.Visible = False: LineOn.Visible = True
CircFlag = 0: CircOff.Visible = True: CircOn.Visible = False
ChkSolidCirc.Enabled = False
BoxFlag = 0: BoxOff.Visible = True: BoxOn.Visible = False
ChkSolidBox.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = True
Label4.Enabled = False
Label5.Enabled = False
End Sub

Private Sub Menu_Click(Index As Integer)
If Index = 0 Then
  Rem New file
  If FileChangedFlag = 1 Then
    a% = MsgBox("Save Current Icon File?", 51, "File NOT Saved")
    If a% = 6 Then
      Rem Yes
      NewAfterSaveFlag = 1
      Call SaveFile
      Exit Sub
    End If
    Rem If a% = 7 Then Stop: Rem No
    If a% = 2 Then Exit Sub: Rem Cancel
  End If
  Call NewFile
  Exit Sub
End If
If Index = 1 Then
  Rem Open file
  Load Form2
  Exit Sub
End If
If Index = 2 Then
  Rem Save
  If Text1.Text = "UNTITLED" Then Load Form3: Exit Sub
  If FileChangedFlag = 1 Then Call SaveFile
  Exit Sub
End If
If Index = 3 Then
  Rem Save As...
  Load Form3
  Exit Sub
End If
If Index = 4 Then Unload Me: Rem Exit
End Sub

Private Sub PenOff_Click()
If CircStartFlag = 1 Then Exit Sub
If LineStartFlag = 1 Then Exit Sub
PenFlag = 1: PenOff.Visible = False: PenOn.Visible = True
FillFlag = 0: FillOff.Visible = True: FillOn.Visible = False
LineFlag = 0: LineOff.Visible = True: LineOn.Visible = False
CircFlag = 0: CircOff.Visible = True: CircOn.Visible = False
ChkSolidCirc.Enabled = False
BoxFlag = 0: BoxOff.Visible = True: BoxOn.Visible = False
ChkSolidBox.Enabled = False
Label1.Enabled = True
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
End Sub


Private Sub VScrollBlue_Change()
If SkipFlag = 0 Then FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True: Call CheckForBlack
b% = 255 - Int(VScrollBlue.Value / 128#)
LblBlueSB.Caption = b%
Pal(SelectedColor) = (Pal(SelectedColor) And 65535) Or (b% * 65536)
Form1.Pall(SelectedColor).BackColor = Pal(SelectedColor)
Call CheckForBlack
End Sub

Private Sub VScrollGreen_Change()
If SkipFlag = 0 Then FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True: Call CheckForBlack
g% = 255 - Int(VScrollGreen.Value / 128#)
LblGreenSB.Caption = g%
Pal(SelectedColor) = (Pal(SelectedColor) And 16711935) Or (g% * 256&)
Form1.Pall(SelectedColor).BackColor = Pal(SelectedColor)
Call CheckForBlack
End Sub

Private Sub VScrollRed_Change()
If SkipFlag = 0 Then FileChangedFlag = 1: CmdRefresh.Enabled = True: TimerRefresh = True: Call CheckForBlack
r% = 255 - Int(VScrollRed.Value / 128#)
LblRedSB = r%
Pal(SelectedColor) = (Pal(SelectedColor) And 16776960) Or r%
Form1.Pall(SelectedColor).BackColor = Pal(SelectedColor)
Call CheckForBlack
End Sub

