VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2775
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   5280
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image6 
         Height          =   855
         Left            =   120
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image5 
         Height          =   855
         Left            =   120
         Top             =   4080
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   855
         Left            =   120
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   120
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   120
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   120
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FigT As String
Public FigInt As Long
Public FigC As String
Private Sub Command1_Click()
'change the type and the image of the figure
Select Case True
Case Option1:
    Form1.FigImg(FigInt).Picture = Image1.Picture
    Form1.FigImg(FigInt).Tag = "q"
Case Option2:
    Form1.FigImg(FigInt).Picture = Image2.Picture
    Form1.FigImg(FigInt).Tag = "o"
Case Option3:
    Form1.FigImg(FigInt).Picture = Image3.Picture
    Form1.FigImg(FigInt).Tag = "h"
Case Option4:
    Form1.FigImg(FigInt).Picture = Image4.Picture
    Form1.FigImg(FigInt).Tag = "r"
Case Option5:
    Form1.FigImg(FigInt).Picture = Image5.Picture
    Form1.FigImg(FigInt).Tag = "p"
Case Option6:
    Form1.FigImg(FigInt).Picture = Image6.Picture
    Form1.FigImg(FigInt).Tag = "k"
End Select

If Form1.IsOnline = True Then
    Form1.WS.SendData "fig:" & FigInt & ":" & FigC & ":" & Form1.FigImg(FigInt).Tag & ":" & Form1.MyName
End If


Unload Me

End Sub
Private Sub Form_Load()
'--------------------
'setting the controls
'--------------------
Me.Caption = "Choose a Figure"
Me.ScaleMode = 3
Frame1.Caption = vbNullString
Command1.Caption = "Ok"


Option1.Caption = "Queen"   'queen
Image1.Picture = LoadPicture(ApPath & "pics\figures\queen_" & FigC & ".ico")


Option2.Caption = "Bishop"  'officer
Image2.Picture = LoadPicture(ApPath & "pics\figures\officer_" & FigC & ".ico")


Option3.Caption = "Knight"  'horse
Image3.Picture = LoadPicture(ApPath & "pics\figures\horse_" & FigC & ".ico")


Option4.Caption = "Rook"    'rock
Image4.Picture = LoadPicture(ApPath & "pics\figures\rock_" & FigC & ".ico")


Option5.Caption = "Pawn"    'pawn
Image5.Picture = LoadPicture(ApPath & "pics\figures\pawn_" & FigC & ".ico")

Option6.Caption = "King"    'king
Image6.Picture = LoadPicture(ApPath & "pics\figures\king_" & FigC & ".ico")

Select Case FigT
    Case "q":
        Option1.Value = True
    Case "o":
        Option2.Value = True
    Case "h":
        Option3.Value = True
    Case "r":
        Option4.Value = True
    Case "p":
        Option5.Value = True
    Case "k":
        Option6.Value = True
End Select

End Sub
