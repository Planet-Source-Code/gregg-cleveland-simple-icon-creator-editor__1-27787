Attribute VB_Name = "VAR_ICN1"
Global BlackPalNum As Integer
Global BoxEndX As Integer
Global BoxEndY As Integer
Global BoxFlag As Byte
Global BoxHeight As Integer
Global BoxStartX As Integer
Global BoxStartY As Integer
Global BoxStartFlag As Byte
Global BoxWidth As Integer
Global BoxX1 As Integer
Global BoxX2 As Integer
Global BoxY1 As Integer
Global BoxY2 As Integer
Global CircEndX As Integer
Global CircEndY As Integer
  Global PrevCircEndX As Integer
  Global PrevCircEndY As Integer
Global CircFlag As Byte
Global CircStartFlag As Byte
Global CircStartX As Integer
Global CircStartY As Integer
Global ClickedPixelX As Integer
Global ClickedPixelY As Integer
Global ExitFlag As Byte
Global FileChangedFlag As Byte
Global FileFoundFlag As Byte
Global FillFlag As Byte
Global FillVals As String
Global Hdr As String
Global LineEndX As Integer
Global LineEndY As Integer
Global LineStartFlag As Byte
Global LineStartX As Integer
Global LineStartY As Integer
Global LineFlag As Byte
Global LoadAfterSaveFlag As Byte
Global LoadName As String
Global NewAfterSaveFlag As Byte
Global Pal(15) As Long
Global PenFlag As Byte
Global PixArray(31, 31) As Long
Global Pos(7) As Integer
Global PreviewX1 As Integer
Global PreviewX2 As Integer
Global PreviewY1 As Integer
Global PreviewY2 As Integer
Global PrevSelectedColor As Integer
Global SaveName As String
Global SelectedColor As Integer
Global SkipFlag As Byte
Global SqWidth As Integer
Global SqHeight As Integer
Global TempPath As String
Global TempColor As Long
Global X As Integer
Global Y As Integer
Sub DrawEditBox()
Form1.Line (BoxX1, BoxY1)-(BoxX2, BoxY2), &H7F7F7F, B
Form1.Line (BoxX1 + 1, BoxY1 + 1)-(BoxX2 - 1, BoxY1 + 1), 0
Form1.Line (BoxX1 + 1, BoxY1 + 1)-(BoxX1 + 1, BoxY2 - 1), 0
Form1.Line (BoxX1 + 1, BoxY2 - 1)-(BoxX2 - 1, BoxY2 - 1), &HDFDFDF
Form1.Line (BoxX2 - 1, BoxY2 - 1)-(BoxX2 - 1, BoxY1 + 0), &HDFDFDF
Form1.Line (BoxX1, BoxY2)-(BoxX2, BoxY2), &HFFFFFF
Form1.Line (BoxX2, BoxY2)-(BoxX2, BoxY1 - 1), &HFFFFFF
End Sub
Sub CheckForBlack()
BlackPalNum = -1
Form1.Img(0).Enabled = True
For n% = 0 To 15
If Pal(n%) = 0 Then BlackPalNum = n%: Exit For
Next n%
If BlackPalNum < (-0.5) Then Form1.Img(0).Checked = False: Form1.Img(0).Enabled = False
End Sub

Sub CopyLittleToBig()
For Y = 0 To 31
For X = 0 To 31
TempColor = Form1.Point(X + PreviewX1 + 2, Y + PreviewY1 + 2)
For c% = 0 To 15
If Pal(c%) = TempColor Then PixArray(X, Y) = c%: Exit For
Next c%
Form1.Line (X * (SqWidth + 2) + BoxX1 + 3, Y * (SqHeight + 2) + BoxY1 + 3)-((X + 1) * (SqWidth + 2) - 2 + BoxX1 + 3, (Y + 1) * (SqHeight + 2) - 2 + BoxY1 + 3), TempColor, BF
Next X
Next Y
FileChangedFlag = 1: Form1.CmdRefresh.Enabled = True: Form1.TimerRefresh = True
Call DrawGrid
End Sub

Sub DrawGrid()
Co& = &HC0C0C0
For X = BoxX1 + 2 To BoxX2 - 2 Step SqWidth + 2
Form1.Line (X, BoxY1 + 2)-(X, BoxY2 - 1), Co&
Next X
For Y = BoxY1 + 2 To BoxY2 - 2 Step SqHeight + 2
Form1.Line (BoxX1 + 2, Y)-(BoxX2 - 1, Y), Co&
Next Y
End Sub

Sub DrawIcon()
For Y = 0 To 31
For X = 0 To 31
    If PixArray(X, Y) = 16 Then
      Form1.Line (X * (SqWidth + 2) + BoxX1 + 3, Y * (SqHeight + 2) + BoxY1 + 3)-((X + 1) * (SqWidth + 2) - 2 + BoxX1 + 3, (Y + 1) * (SqHeight + 2) - 2 + BoxY1 + 3), Pal(SelectedColor), BF
      Form1.PSet (X + PreviewX1 + 2, Y + PreviewY1 + 2), Pal(SelectedColor)
    End If
    If PixArray(X, Y) < 16 Then
Form1.Line (X * (SqWidth + 2) + BoxX1 + 3, Y * (SqHeight + 2) + BoxY1 + 3)-((X + 1) * (SqWidth + 2) - 2 + BoxX1 + 3, (Y + 1) * (SqHeight + 2) - 2 + BoxY1 + 3), Pal(PixArray(X, Y)), BF
Form1.PSet (X + PreviewX1 + 2, Y + PreviewY1 + 2), Pal(PixArray(X, Y))
    End If
Next X
Next Y
End Sub

Sub DrawPreviewBox()
Form1.Line (PreviewX1, PreviewY1)-(PreviewX2, PreviewY2), &H7F7F7F, B
Form1.Line (PreviewX1 + 1, PreviewY1 + 1)-(PreviewX2 - 1, PreviewY1 + 1), 0
Form1.Line (PreviewX1 + 1, PreviewY1 + 1)-(PreviewX1 + 1, PreviewY2 - 1), 0
Form1.Line (PreviewX1 + 1, PreviewY2 - 1)-(PreviewX2 - 1, PreviewY2 - 1), &HDFDFDF
Form1.Line (PreviewX2 - 1, PreviewY2 - 1)-(PreviewX2 - 1, PreviewY1 + 0), &HDFDFDF
Form1.Line (PreviewX1, PreviewY2)-(PreviewX2, PreviewY2), &HFFFFFF
Form1.Line (PreviewX2, PreviewY2)-(PreviewX2, PreviewY1 - 1), &HFFFFFF
End Sub

Sub LoadIconFile()
If LoadAfterSaveFlag = 1 Then GoTo SkipCheck
If FileChangedFlag = 1 Then
  a% = MsgBox("Save Current Icon File?", 52, "File NOT Saved")
  If a% = 6 Then
    Rem Yes
    LoadAfterSaveFlag = 1
    Call SaveFile
    Exit Sub
  End If
  Rem If a% = 7 Then Stop: Rem No
  If a% = 2 Then Exit Sub: Rem Cancel
End If
SkipCheck:
LoadOffset% = 0
If Right$(LoadName, 4) = ".bmp" Then
  Open LoadName For Binary As #1
  Get #1, 19, w%
  Get #1, 23, h%
  Get #1, 29, b%
  If (w% <> 32) Or (h% <> 32) Or (b% <> 4) Then
    Close
    X% = MsgBox("Oops!  Must be 32X32, 16 Colors!", 48, "Wrong Format...")
    Exit Sub
  End If
  Close
  LoadOffset% = 8
End If
Open LoadName For Binary As #1
For a% = 0 To 15
Get #1, a% * 4 + 63 - LoadOffset%, Pal(a%)
r% = Int(Pal(a%) / 65536) And 255
g% = Int(Pal(a%) / 256) And 255
b% = Pal(a%) And 255
Pal(a%) = b% * 65536 + g% * 256& + r%
Next a%
For Y = 0 To 31
For X = 0 To 30 Step 2
Get #1, Y * 16 + 127 + X / 2 - LoadOffset%, D%: D% = D% And 255
PixArray(X, 31 - Y) = ((D% And 240) / 16)
PixArray(X + 1, 31 - Y) = (D% And 15)
Next X
Next Y
Close
Call DrawIcon
Call ShowPalette
FileChangedFlag = 0: Form1.CmdRefresh.Enabled = False: Form1.TimerRefresh = False
LoadAfterSaveFlag = 0
SaveName = LoadName
Form1.Text1.Text = LoadName
End Sub

Sub NewFile()
NewAfterSaveFlag = 0
FileFoundFlag = 0
FileChangedFlag = 0: Form1.CmdRefresh.Enabled = False: Form1.TimerRefresh = False
SaveName = "UNTITLED"
Form1.Text1.Text = "UNTITLED"
Pal(0) = 0
Pal(1) = 8323072
Pal(2) = 32512&
Pal(3) = 8355584
Pal(4) = 127
Pal(5) = 8323199
Pal(6) = 32639&
Pal(7) = 12632256
Pal(8) = 8355711
Pal(9) = 16711680
Pal(10) = 65280
Pal(11) = 16776960
Pal(12) = 255
Pal(13) = 16711935
Pal(14) = 65535
Pal(15) = 16777215
Call ShowPalette
For X = 0 To 31
For Y = 0 To 31
PixArray(X, Y) = 15
Next Y
Next X
Form1.Line (BoxX1 + 2, BoxY1 + 2)-(32 * (SqWidth + 2) + BoxX1 + 2, 32 * (SqHeight + 2) + BoxY1 + 2), 16777215, BF
Form1.Line (PreviewX1 + 2, PreviewY1 + 2)-(PreviewX2 - 2, PreviewY2 - 2), 16777215, BF
Call DrawGrid
End Sub

Sub OutlineBox()
Rem If Form1.Pall(SelectedColor).Visible = False Then Exit Sub
Rem Un-outline first!
Form1.Line (Form1.Pall(PrevSelectedColor).Left - 2, Form1.Pall(PrevSelectedColor).Top - 2)-(Form1.Pall(PrevSelectedColor).Left + Form1.Pall(PrevSelectedColor).Width + 1, Form1.Pall(PrevSelectedColor).Top + Form1.Pall(PrevSelectedColor).Height + 1), &HC0C0C0, B
Form1.FillColor = Pal(SelectedColor)
Form1.Line (Form1.Pall(SelectedColor).Left - 2, Form1.Pall(SelectedColor).Top - 2)-(Form1.Pall(SelectedColor).Left + Form1.Pall(SelectedColor).Width + 1, Form1.Pall(SelectedColor).Top + Form1.Pall(SelectedColor).Height + 1), 0, B
PrevSelectedColor = SelectedColor
b% = Int(Pal(SelectedColor) / 65536) And 255
g% = Int(Pal(SelectedColor) / 256) And 255
r% = Pal(SelectedColor) And 255
Rem Pal(a%) = r% * 65536 + g% * 256& + b%
SkipFlag = 1: Rem Prevents FileChangeFlag and CmdRefresh from being enabled
Form1.VScrollBlue.Value = 32767 - b% * 128
Form1.VScrollGreen.Value = 32767 - g% * 128
Form1.VScrollRed.Value = 32767 - r% * 128
SkipFlag = 0: Rem Allow FileChangeFlag and CmdRefresh to being enabled
End Sub

Sub PenOff()
If CircStartFlag = 1 Then Exit Sub
If LineStartFlag = 1 Then Exit Sub
PenFlag = 1: Form1.PenOff.Visible = False: Form1.PenOn.Visible = True
FillFlag = 0: Form1.FillOff.Visible = True: Form1.FillOn.Visible = False
LineFlag = 0: Form1.LineOff.Visible = True: Form1.LineOn.Visible = False
CircFlag = 0: Form1.CircOff.Visible = True: Form1.CircOn.Visible = False
Form1.ChkSolidCirc.Enabled = False
BoxFlag = 0: Form1.BoxOff.Visible = True: Form1.BoxOn.Visible = False
Form1.ChkSolidBox.Enabled = False
Form1.Label1.Enabled = True
Form1.Label2.Enabled = False
Form1.Label3.Enabled = False
Form1.Label4.Enabled = False
Form1.Label5.Enabled = False
End Sub

Sub SaveFile()
If SaveName = "UNTITLED" Then Load Form3: Exit Sub
Rem SaveName = LoadName
If (GetAttr(SaveName) And 1) = 1 Then
  a% = MsgBox("File is Read-Only", 48, "Save Error")
  LoadAfterSaveFlag = 0
  ExitFlag = 0
  NewAfterSaveFlag = 0
  Exit Sub
End If
Form1.MousePointer = 11
Open SaveName For Binary As #1
For X = 1 To 62
D$ = Chr$(Val("&H" + Mid$(Hdr, X * 2 - 1, 2)))
Put #1, X, D$
Next X
For a% = 0 To 15
b% = Int(Pal(a%) / 65536) And 255
g% = Int(Pal(a%) / 256) And 255
r% = Pal(a%) And 255
p& = r% * 65536 + g% * 256& + b%
Put #1, a% * 4 + 63, p&
Next a%
For Y = 0 To 31
For X = 0 To 30 Step 2
D$ = Chr$(PixArray(X, 31 - Y) * 16 Or PixArray(X + 1, 31 - Y))
Put #1, Y * 16 + 127 + X / 2, D$
Next X
Next Y
If Form1.Img(0).Checked = True Then
  Rem Transparent black
  For Y = 0 To 31
  For X = 0 To 24 Step 8
  Sum% = 0
  Byt% = 639 + Int((X + (31 - Y) * 32) / 8)
  For X2% = 0 To 7
  Bit% = 0
  If PixArray(X + X2%, Y) = BlackPalNum Then Bit% = 1
  Rem If PixArray(X + X2%, Y) = 0 Then Bit% = 1
  Sum% = Sum% Or (Pos(X2%) * Bit%)
  Next X2%
  D$ = Chr$(Sum%)
  Put #1, Byt%, D$
  Next X
  Next Y
End If
If Form1.Img(0).Checked = False Then
  Rem Must pad last 128 bytes with zeros.
  D$ = String$(128, 0)
  Put #1, 639, D$
End If
Close
Form1.MousePointer = 1
If ExitFlag = 1 Then Unload Form1: End
FileChangedFlag = 0: Form1.CmdRefresh.Enabled = False: Form1.TimerRefresh = False
If LoadAfterSaveFlag = 1 Then Call LoadIconFile
Form1.Text1.Text = SaveName
If NewAfterSaveFlag = 1 Then Call NewFile
End Sub
Sub ShowPalette()
For a% = 0 To 15
Form1.Pall(a%).Visible = True
Form1.Pall(a%).BackColor = Pal(a%)
Next a%

SelectedColor = 0
Call OutlineBox
Call PenOff
Form1.LblPal.Visible = True
End Sub

