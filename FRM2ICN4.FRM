VERSION 4.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Open File:"
   ClientHeight    =   3180
   ClientLeft      =   2016
   ClientTop       =   2580
   ClientWidth     =   7260
   ControlBox      =   0   'False
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
   Height          =   3504
   Icon            =   "FRM2ICN4.frx":0000
   Left            =   1968
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7260
   Top             =   2304
   Width           =   7356
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   2412
   End
   Begin VB.CommandButton CmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   372
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1944
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2412
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   2412
   End
   Begin VB.Label LabelDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drives:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label LabelDirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Directories:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label LabelFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File &Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Rem  List, ListCount, and ListIndex

Private Sub CmdCancel_Click()
Unload Form2
Form1.Enabled = True
Form1.SetFocus
End Sub

Private Sub CmdOK_Click()
On Error GoTo ErrHandler
If RTrim$(LTrim$(Text1.Text)) = "\" Then
  Dir1.Path = Left$(Drive1.Drive, 2) + "\"
  Exit Sub
End If
AsteriskFlag% = 0
For a% = 1 To Len(Text1.Text)
If Mid$(Text1.Text, a%, 1) = "*" Then AsteriskFlag% = 1
Next a%
If AsteriskFlag% = 1 Then
  Rem Is it a pattern?
  File1.Pattern = Text1.Text
  GoTo FocusThenExit
End If
Rem if no asterisk in Text1.Text1 then...
Rem Is it a file name?
LoadName = File1.Path
If Right$(LoadName, 1) <> "\" Then LoadName = LoadName + "\"
LoadName = LoadName + Text1.Text
Rem If DotFlag = 0 Then LoadName = LoadName + ".ico"
E% = 0
If Len(Dir$(LoadName, 6)) > 0 Then GoTo Done

'Is it a directory?
t$ = Text1.Text
If Right$(t$, 1) = "\" Then t$ = Left$(t$, Len(t$) - 1)
If Left$(t$, 1) = "\" Then
  t$ = Left$(Drive1.Drive, 2) + t$
  GoTo TestDir
End If

Rem t$ = Text1.Text
If Right$(Dir1.Path, 1) <> "\" Then t$ = "\" + t$
t$ = Dir1.Path + t$
E% = 0
TestDir:
E% = 0
Dir1.Path = t$
If E% = 0 Then GoTo FocusThenExit
a% = MsgBox("Invalid File/Directory", 48, "Error")
GoTo FocusThenExit

Done:
TempDrive = Drive1.Drive
TempPath = Dir1.Path
TempIndex = File1.ListIndex
Call GotFileName
Unload Form2
Form1.Enabled = True
Form1.SetFocus
Call LoadIconFile
Exit Sub

ErrHandler:
E% = Err
Resume Next

FocusThenExit:
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
Rem Text1.Text = File1.FileName
Rem  List, ListCount, and ListIndex
If File1.ListCount > 0 Then
  Text1.Text = File1.List(0)
End If
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrHandler
Dir1.Path = Drive1.Drive
Exit Sub
ErrHandler:
Resume Next
End Sub

Private Sub File1_Click()
Text1.Text = File1.filename
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub File1_DblClick()
LoadName = File1.Path
If Right$(LoadName, 1) <> "\" Then LoadName = LoadName + "\"
LoadName = LoadName + File1.List(File1.ListIndex)
TempDrive = Drive1.Drive
TempPath = Dir1.Path
TempIndex = File1.ListIndex
Rem Form2.Enabled = False
Rem Form2.Visible = False
Call GotFileName
Unload Form2
Form1.Enabled = True
Form1.SetFocus
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Rem Form2.Enabled = False
  Rem Form2.Visible = False
  Call CmdOK_Click
End If
End Sub

Private Sub File1_PathChange()
Text1.Text = File1.filename
End Sub

Private Sub File1_PatternChange()
Rem  List, ListCount, and ListIndex
If File1.ListCount > 0 Then Text1.Text = File1.List(0)
End Sub

Private Sub Form_Load()
FileFoundFlag = 0
Form1.Enabled = False
File1.Pattern = "*.ico"
Width = 7356
Height = 3504
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
LabelFileName.Left = 240
LabelFileName.Top = 120
LabelFileName.Width = 1572
LabelFileName.Height = 252
LabelDirectory.Left = 3000
LabelDirectory.Top = 120
LabelDirectory.Width = 1212
LabelDirectory.Height = 192
LabelDrives.Left = 3000
LabelDrives.Top = 2520
LabelDrives.Width = 1212
LabelDrives.Height = 192
Text1.Left = 240
Text1.Top = 360
Text1.Width = 2412
Text1.Height = 288
File1.Left = 240
File1.Top = 720
File1.Width = 2412
File1.Height = 1944
Dir1.Left = 3000
Dir1.Top = 360
Dir1.Width = 2415
Dir1.Height = 2055
Drive1.Left = 3000
Drive1.Top = 2760
Drive1.Width = 2412
Rem Drive1.Height = 288 (Read only!)
CmdOK.Left = 5640
CmdOK.Top = 120
CmdOK.Width = 1452
CmdOK.Height = 372
CmdCancel.Left = 5640
CmdCancel.Top = 720
CmdCancel.Width = 1452
CmdCancel.Height = 372
Visible = True
If TempDrive = "" Then
  Open "c:\iconpath.cfg" For Binary As #1
  Close
  Open "c:\IconPath.cfg" For Input As #1
  l& = LOF(1)
  If l& > 0 Then Line Input #1, p$
  Close
  If l& > 0 Then
    TempDrive = Left$(p$, 2)
    TempPath = Mid$(p$, 3)
  End If
End If
Rem File1.ListIndex = 0
If TempDrive > "" Then
  Drive1.Drive = TempDrive
  Dir1.Path = TempPath
  Rem If File1.ListCount > 0 Then File1.ListIndex = TempIndex
End If
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub GotFileName()
Tmp$ = TempPath + Chr$(13) + Chr$(10)
Open "c:\IconPath.cfg" For Binary As #1
Put #1, 1, Tmp$
Close
FileFoundFlag = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CmdOK_Click
End If
End Sub

