VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form6.frx":2C5A
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   3480
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'send START to the other player
Form1.WS.SendData "sta"
'show the board
Unload Me
Form1.Show
Form1.StartMulGame

End Sub
Private Sub Command2_Click()
'return to the main menu
Form1.RtnMain

End Sub
Private Sub Form_Load()
'setting the controls
Me.ScaleMode = 3
Me.Caption = App.Title & " Main Lobby"

Image1.Width = 137
Image1.Height = 135
Image1.Picture = LoadPicture(ApPath & "pics\swords.bmp")

Label5.Caption = "Players:"
Label6.Caption = "Color:"

Text1.Text = "-- Server Ready."
Text2.Text = vbNullString

Text2.Top = Image1.Top + Image1.Height - Text2.Height

Command2.Caption = "Exit"
Command1.Caption = "Start"

'check if i am the server
If Form1.IsServer = True Then
    Command1.Enabled = True
    'restrict changing the client's color
    Label3.Enabled = True
    Label4.Enabled = False
Else
    Command1.Enabled = False
    'restrict changing the server's color
    Label3.Enabled = False
    Label4.Enabled = True
End If

Label3.Caption = vbNullString
Label4.Caption = vbNullString
Label3.BackColor = &HFFFFFF
Label4.BackColor = &H0&

Timer1.Interval = 1
Timer1.Enabled = True

End Sub
Private Sub ChColor(LabelName As Label)
'toggle the label's colors
If LabelName.BackColor = &HFFFFFF Then
    LabelName.BackColor = &H0&
Else
    LabelName.BackColor = &HFFFFFF
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
'stop the timer
Timer1.Enabled = False
End Sub

Private Sub Label3_Click()
'change the color
ChColor Label3

'send my color choise to the other player
Form1.WS.SendData "clr:" & Label3.BackColor

End Sub
Private Sub Label4_Click()
'change the color
ChColor Label4

'send my color choise to the other player
Form1.WS.SendData "clr:" & Label4.BackColor

End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'use enter to send message
If KeyCode = vbKeyReturn Then
    Text1.Text = Text1.Text & vbNewLine & Form1.MyName & ">>" & Text2.Text
    Form1.WS.SendData "lms:" & Form1.MyName & ">>" & Text2.Text
    Text2.Text = vbNullString
End If

End Sub
Private Sub Timer1_Timer()
'check if the two player colors are the same
If Form1.IsServer = True Then
    If Label3.BackColor = Label4.BackColor Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End If
'determine my color
If Form1.IsServer = True Then   'look at label3
    If Label3.BackColor = &HFFFFFF Then  'is it white?
        Form1.MyColor = "w" 'Yep
    Else
        Form1.MyColor = "b" 'Nope
    End If
Else    'look at label4
    If Label4.BackColor = &HFFFFFF Then  'is it white?
        Form1.MyColor = "w" 'Yep
    Else
        Form1.MyColor = "b" 'Nope
    End If
End If

End Sub
