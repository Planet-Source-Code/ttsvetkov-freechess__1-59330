VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   2115
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   1940
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text1Def As String  'the default text in text1
Private Sub Command1_Click()
'go to singleplayer
Form1.IsOnline = False
Form1.Label4.Caption = vbNullString
Form1.Command5.Enabled = False

Form1.Show
Unload Me

End Sub
Private Sub Command2_Click()
'join
Form1.IsServer = False

Form1.WS.LocalPort = 0
Form1.WS.RemotePort = 5559
Form1.WS.RemoteHost = Text1.Text
Form1.WS.Connect

End Sub
Private Sub Command3_Click()
'Form6.Show
'Unload Me
Form1.IsServer = True

'start the ws listening
Form1.WS.RemotePort = 0
Form1.WS.LocalPort = 5559
Form1.WS.Listen
Label1.Caption = "WinSock Listening Started." & vbNewLine & Form1.WS.LocalIP & " (" & Form1.WS.LocalHostName & ")"
'disable clicking on other controls
Command3.Enabled = False
Command4.Enabled = True
Command6.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False

End Sub
Private Sub Command4_Click()
'stop the listening
Form1.WS.CloseSck
Label1.Caption = "WinSock Closed."
'reenable the controls
Command3.Enabled = True
Command4.Enabled = False
Command6.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True


End Sub
Private Sub Command5_Click()
'------------------
'exit from the game
'------------------
Form1.WS.CloseSck
Unload Form1
Unload Form2
Unload Form3
Unload Form4

End Sub
Private Sub Command6_Click()
'show the about
Form1.menu31_Click 0

End Sub
Private Sub Form_Load()
'----------------------------------------
'checking for another instance of the app
'----------------------------------------
If App.PrevInstance = True Then
    MsgBox App.Title & " is already running.", vbInformation, App.Title
    End
End If
'--------------------
'setting the controls
'--------------------
Text1Def = "Enter the host's IP or Name here..."

Me.Caption = App.Title
Me.ScaleMode = 3

Option1.Caption = "Singleplayer Training (no AI, sorry :)"
Option2.Caption = "Join a Multiplayer game"
Option3.Caption = "Start a Multiplayer game"

Command1.Caption = "Go"
Command2.Caption = "Join"
Command3.Caption = "Start"
Command4.Caption = "Stop"
Command5.Caption = "Exit"
Command6.Caption = "About"

Frame1.Caption = " Main Menu "

Text1.Text = Text1Def

Label1.Caption = "Idle."

Option1.Value = True
'show the main menu picture
Picture1.ScaleMode = 3
Picture1.Picture = LoadPicture(ApPath & "pics\menu.bmp")
Picture1.Width = 145
Picture1.Height = 193

'load form1
Load Form1
Form1.Hide

End Sub
Private Sub Option1_Click()
CtlToggle
End Sub
Private Sub Option2_Click()
CtlToggle
End Sub
Private Sub Option3_Click()
CtlToggle
End Sub
Private Sub CtlToggle()
'-------------------
'toggle the controls
'-------------------
Command1.Enabled = Option1.Value

Text1.Enabled = Option2.Value
Command2.Enabled = Option2.Value

Command3.Enabled = Option3.Value
Command4.Enabled = False

'return the default text if nothing has been entered
If Text1.Text = vbNullString Then
    Text1.Text = Text1Def
End If

End Sub
Private Sub Text1_Click()
'erase the text inside the textbox
If Text1.Text = Text1Def Then
    Text1.Text = vbNullString
End If

End Sub
