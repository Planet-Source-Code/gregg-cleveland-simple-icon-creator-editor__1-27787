VERSION 4.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Save As..."
   ClientHeight    =   3192
   ClientLeft      =   1044
   ClientTop       =   3192
   ClientWidth     =   7260
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
   Height          =   3516
   Icon            =   "FRM3ICN4.frx":0000
   Left            =   996
   LinkTopic       =   "Form3"
   ScaleHeight     =   3192
   ScaleWidth      =   7260
   Top             =   2916
   Width           =   7356
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   2412
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1968
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   2412
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1944
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2412
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   372
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1452
   End
   Begin VB.CommandButton CmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
      Width           =   2412
   End
   Begin VB.Label LabelFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File &Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1572
   End
   Begin VB.Label LabelDirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Directories:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label LabelDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drives:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   1212
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
Unload Form3
Form1.Enabled = True
If Form1.Visible = True Then Form1.SetFocus
ExitFlag = 0
End Sub

Private Sub CmdOK_Click()
On Error GoTo ErrHandler
If Text1.Text = "" Then
  Text1.Text = "*.ico"
  File1.Pattern = Text1.Text
  Exit Sub
End If
AsteriskFlag% = 0
DotFlag% = 0
For a% = 1 To Len(Text1.Text)
If Mid$(Text1.Text, a%, 1) = "*" Then AsteriskFlag% = 1
If Mid$(Text1.Text, a%, 1) = "." Then DotFlag% = 1
Next a%
Rem if no asterisk in Text1.Text1 then...
If AsteriskFlag% = 0 Then
  Rem Is it a path name?
  If Right$(Text1.Text, 1) = "\" Or Left$(Text1.Text, 1) = "\" Then
    If Left$(Text1.Text, 1) <> "\" Then Text1.Text = "\" + Text1.Text
    If Right$(Text1.Text, 1) = "\" Then Text1.Text = Left$(Text1.Text, Len(Text1.Text) - 1)
    Dir1.Path = Text1.Text
    Text1.Text = "*.ico"
    File1.Pattern = Text1.Text
    Exit Sub
  End If
  Rem Must be file name or pattern!
End If
If AsteriskFlag% = 1 Then
  Rem Is it a pattern?
  File1.Pattern = Text1.Text
  Exit Sub
End If
Rem Not a pattern, must be a file name!
If DotFlag% = 0 Then Text1.Text = Text1.Text + ".ico"
If DotFlag% > 1 Then Text1.Text = "Not valid!": Exit Sub
SaveName = File1.Path
If Right$(SaveName, 1) <> "\" Then SaveName = SaveName + "\"
SaveName = SaveName + Text1.Text
E% = 0
Open SaveName For Input As #1
Rem L& = LOF(1)
Close
If (E% > 0) And (E% <> 53) Then
  Rem Must except non-existent name!
  Text1.Text = "Not valid!"
  Exit Sub
End If
Rem A valid file name...
Open SaveName For Binary As #1
If LOF(1) > 0 Then
  Close
  a% = MsgBox("Overwrite Existing File?", 308, "File Exists")
  If a% = 6 Then GoTo Skip1: Rem Yes
  If a% = 7 Then
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Exit Sub: Rem No
  End If
End If
Skip1:
Close
If ExitFlag = 1 Then
  Unload Form3
  Call SaveFile
  If ExitFlag = 1 Then
    Unload Form1
    End
  End If
  If ExitFlag = 0 Then
    Form1.Enabled = True
    Form1.SetFocus
    Exit Sub
  End If
End If
TempDrive = Drive1.Drive
TempIndex = File1.ListIndex
If TempPath <> Dir1.Path Then
  TempPath = Dir1.Path
  Call GotFileName
End If
Form1.Enabled = True
FileFoundFlag = 0
Rem If Form1.Text1.Text = "UNTITLED" Then Form1.Text1.Text = SaveName
Form1.SetFocus
Call SaveFile
Unload Form3
Exit Sub

ErrHandler:
E% = Err
Resume Next
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Rem Text1.Text = File1.FileName
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text1.Text = File1.filename
End Sub

Private Sub File1_DblClick()
SaveName = File1.Path
If Right$(SaveName, 1) <> "\" Then SaveName = SaveName + "\"
SaveName = SaveName + File1.List(File1.ListIndex)
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

Private Sub Form_Load()
FileFoundFlag = 0
Form1.Enabled = False
Drive1.Drive = "c:"
Dir1.Path = "\"
File1.Pattern = "*.ico"
t$ = ""
For a% = Len(LoadName) To 1 Step -1
t2$ = Mid$(LoadName, a%, 1)
If t2$ = "\" Then Exit For
t$ = t2$ + t$
Next a%

DotFlag% = 0
For a% = 1 To Len(t$)
If Mid$(t$, a%, 1) = "." Then
  t$ = Left$(t$, a%) + "ico"
  DotFlag% = 1
  Exit For
End If
Next a%
If DotFlag% = 0 Then t$ = Left$(t$, a%) + ".ico"
Text1.Text = t$

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
Width = 7356
Height = 3504
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

