VERSION 4.00
Begin VB.Form Form4 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2004
   ClientLeft      =   3480
   ClientTop       =   3180
   ClientWidth     =   5772
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2328
   Icon            =   "Frm4icn4.frx":0000
   Left            =   3432
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   Top             =   2904
   Width           =   5868
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Converts 16 color bitmaps (32X32) to  icons"
      Height          =   252
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   4812
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Edits  Icons"
      Height          =   252
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   4212
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Creates  Icons"
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Icon Editor written by Gregg Cleveland"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4932
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
Unload Form4
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
Visible = True
End Sub


