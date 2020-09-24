VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   2895
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'exit
Unload Me
End Sub

Private Sub Form_Load()
'setting the controls
Me.ScaleMode = 3
Picture1.ScaleMode = 3
Picture1.Height = 214
Picture1.Width = 160
Picture1.Picture = LoadPicture(ApPath & "pics\about.bmp")

'setting the text in the form
Command1.Caption = "Exit"
Label1.Caption = _
App.Title & vbNewLine & App.Major & "." & App.Minor & "." & App.Revision & " Beta"

Label2.Caption = _
"A new PC version of the popular game. " & App.Title & " is designed so, the players have free control over the figures on the board, thus making it excellent for learning the rules of playing chess. You can move the figures whatever you want to, replace them, change them. You can also play for fun with your buddies via Internet or LAN. For now, no AI has been implemented, due to the big difficulties in creating enoug intelligent and logical AI." _
& vbNewLine & vbNewLine & "2 days of graphical design." & vbNewLine & "3 days of network coding." & vbNewLine & "5 days of main coding." & vbNewLine & "Written in VisualBasic 6.0. Requires the Microsoft Common Dialog control (Comdlg32.ocx)."

Label3.Caption = _
"Written by Todor Tzvetkov" & vbNewLine & "snake@contact.bg" & vbNewLine & "13.2.2005"
End Sub
