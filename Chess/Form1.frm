VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11475
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   7680
      Width           =   2535
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   7680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Form1.frx":2C5A
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9120
      Top             =   8520
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   360
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   10800
      Begin VB.Image MovBox 
         Height          =   735
         Index           =   0
         Left            =   9120
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label3"
         Height          =   1335
         Left            =   8760
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   8760
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   2775
         Left            =   6120
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Image FigImg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   780
         Index           =   0
         Left            =   8160
         Top             =   120
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         Height          =   735
         Left            =   7320
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   7680
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label CrdLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "CrdLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   8160
      TabIndex        =   4
      Top             =   8640
      Width           =   900
   End
   Begin VB.Menu menu1 
      Caption         =   "Game"
      Index           =   0
      Begin VB.Menu menu11 
         Caption         =   "New"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu menu12 
         Caption         =   "Exit"
         Index           =   1
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Options"
      Index           =   1
      Begin VB.Menu menu21 
         Caption         =   "Settings"
         Index           =   0
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu menu3 
      Caption         =   "Help"
      Index           =   2
      Begin VB.Menu menu31 
         Caption         =   "About"
         Index           =   0
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu popup1 
      Caption         =   "popup1"
      Visible         =   0   'False
      Begin VB.Menu popup11 
         Caption         =   "Change the figure..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the sound API
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'the winsocks
Public WithEvents WS As CSocketMaster
Attribute WS.VB_VarHelpID = -1

'the APPATH
Dim ApPath As String

'the boxes
Private Type BoxProps
    XStart As Long
    YStart As Long
    IsSafe As Boolean
    ImgIndex As Long
End Type
Dim Box(1 To 96) As BoxProps

'the data buffer for the winsock
Dim Dat As String

'"from" box number
Dim FromBoxIn As Long

'the IsOnline swich
Public IsOnline As Boolean

'the figure's moving speed
Dim MovSp As Long

'my color
Public MyColor As String

'my turn
Dim MyTurn As Boolean

'if i am the server
Public IsServer As Boolean

'the other player's name
Dim YourName As String

'options
Public FigMovSp As Long 'figure's moving speed in singleplay
Public SelColor As Long 'the color of the selection box
Public ShowMov As Boolean   'the movement colors swich
Public MyName As String 'my name :)
Public Snd As Boolean   'the sound swich
Dim OnlineFigMovSp As Long  'the moving speed during multiplay
Public EndTurn As Boolean


'the figure smooth moving
Dim FigIn As Long
Dim ToBox As Long

'the safeboxes
Dim SBox(1 To 32) As Long
Private Sub Command1_Click()
'send a message to the other player
Text1.Text = Text1.Text & vbNewLine & MyName & ">>" & Text2.Text
If IsOnline = True Then
    WS.SendData "msg:" & MyName & ">>" & Text2.Text
End If
Text2.Text = vbNullString
Text2.SetFocus

End Sub
Private Sub Command2_Click()
'exit the game
Unload Me

End Sub
Private Sub Command3_Click()
'play the check sound
If Snd = True Then
    PlaySound ApPath & "sound\check.wav", vbEmpty, SND_FILENAME & SND_ASYNC
End If
Text1.Text = Text1.Text & vbNewLine & "-- " & MyName & " says CHECK!"
'send the notification to the other player
If IsOnline = True Then
    WS.SendData "chk:" & MyName
End If

End Sub
Private Sub Command4_Click()
'play the castling sound
If Snd = True Then
    PlaySound ApPath & "sound\castling.wav", vbEmpty, SND_FILENAME & SND_ASYNC
End If
Text1.Text = Text1.Text & vbNewLine & "-- " & MyName & " says CASTLING!"
'send the notification to the other player
If IsOnline = True Then
    WS.SendData "csl:" & MyName
End If

End Sub
Private Sub Command5_Click()
'--------
'end turn
'--------
'send ENDTURN to the other player
WS.SendData "end"

'disable my controls
Picture1.Enabled = False
Command5.Enabled = False

'disable the notifications
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False

'update the color in the colorbox
If MyColor = "w" Then
    Label4.ForeColor = &H0&
Else
    Label4.ForeColor = &HFFFFFF
End If
Label4.Caption = YourName

MyTurn = False

End Sub
Private Sub Command6_Click()
'play the checkmate sound
If Snd = True Then
    PlaySound ApPath & "sound\checkmate.wav", vbEmpty, SND_FILENAME & SND_ASYNC
End If
Text1.Text = Text1.Text & vbNewLine & "-- " & MyName & " says CHECKMATE!"
'send the notification to the other player
If IsOnline = True Then
    WS.SendData "chm:" & MyName
End If

End Sub

Private Sub Form_Load()
On Error GoTo Err
'---------------------------
'loading the winsock objects
'---------------------------
Set WS = New CSocketMaster
'------------------
'finding the appath
'------------------
If Right(App.Path, 1) = "\" Then
    ApPath = App.Path
Else
    ApPath = App.Path & "\"
End If
'-------------------------
'setting the objects props
'-------------------------
Me.ScaleMode = 3
Me.Caption = App.Title

Picture1.ScaleMode = 3
Picture1.Width = 8 * 60 + (4 * 60)
Picture1.Height = 8 * 60
Picture1.Picture = LoadPicture(ApPath & "pics\wboard.bmp")

FigImg(0).Visible = False
FigImg(0).Enabled = False
FigImg(0).Width = 60
FigImg(0).Height = 60

MovBox(0).Enabled = False
MovBox(0).Visible = False

Label1.Enabled = False
Label1.Top = 0
Label1.Left = 0
Label1.Width = 8 * 60
Label1.Height = 8 * 60
Label1.BackStyle = 0
Label1.BorderStyle = 1
Label1.Caption = vbNullString
Label1.ZOrder 1

Label2.Enabled = False
Label2.Top = 0
Label2.Left = 8 * 60
Label2.Width = 4 * 60
Label2.Height = 4 * 60
Label2.BackStyle = 1
Label2.BorderStyle = 1
Label2.Caption = vbNullString
Label2.ZOrder 1

Label3.Enabled = False
Label3.Top = 4 * 60
Label3.Left = 8 * 60
Label3.Width = 4 * 60
Label3.Height = 4 * 60
Label3.BackStyle = 1
Label3.BorderStyle = 1
Label3.Caption = vbNullString
Label3.ZOrder 1

Shape1.Height = 60
Shape1.Width = 60
Shape1.Visible = False
Shape1.ZOrder 0

Timer1.Enabled = False
Timer1.Interval = 1

Text1.Text = "-- Game Initialized."
Text2.Text = vbNullString
Command1.Caption = "Send"

Command2.Caption = "Exit"
Command3.Caption = "Check!"
Command5.Caption = "End Turn"
Command6.Caption = "Checkmate!"
Frame1.Caption = "Notifications"
Command4.Caption = "Castling!"

WS.Protocol = sckTCPProtocol
WS.LocalPort = 5559
WS.RemotePort = 5559
'---------------------------------------------
'setting some default values for the variables
'---------------------------------------------
FromBoxIn = 1
FigMovSp = 20
SelColor = Shape1.BorderColor
ShowMov = True
IsOnline = False
MyName = "Unnamed"
Snd = True
OnlineFigMovSp = 10
EndTurn = True
'--------------------------------------
'setting the box numbers and dimentions
'--------------------------------------
Dim RowNum As Long
Dim BoxNum As Long
Dim x As Long
Dim y As Long
x = 0
y = 0
For RowNum = 0 To 7 Step 1
    For BoxNum = (RowNum * 12 + 1) To (RowNum * 12 + 12) Step 1
        Box(BoxNum).XStart = x
        x = x + 60
        Box(BoxNum).YStart = y
        Box(BoxNum).ImgIndex = 0
        Box(BoxNum).IsSafe = False
    Next
    x = 0
    y = (RowNum + 1) * 60
Next
'--------------------------
'setting the safebox nubers
'--------------------------
Dim Num As Long
'row 1
For Num = 1 To 4 Step 1
    SBox(Num) = (1 * 8) + Num
Next
'row 2
For Num = 5 To 8 Step 1
    SBox(Num) = (2 * 8) + Num
Next
'row 3
For Num = 9 To 12 Step 1
    SBox(Num) = (3 * 8) + Num
Next
'row 4
For Num = 13 To 16 Step 1
    SBox(Num) = (4 * 8) + Num
Next
'row 5
For Num = 17 To 20 Step 1
    SBox(Num) = (5 * 8) + Num
Next
'row 6
For Num = 21 To 24 Step 1
    SBox(Num) = (6 * 8) + Num
Next
'row 7
For Num = 25 To 28 Step 1
    SBox(Num) = (7 * 8) + Num
Next
'row 8
For Num = 29 To 32 Step 1
    SBox(Num) = (8 * 8) + Num
Next
'-----------------------------
'setting which boxes are safes
'-----------------------------
For Num = 1 To UBound(SBox) Step 1
    Box(SBox(Num)).IsSafe = True
Next
'-----------------------------------------------
'setting the figure images and positions
'from 1 to 16 are black, from 17 to 32 are white
'-----------------------------------------------
'placing the black pawns (1 - 8)
For Num = 1 To 8 Step 1
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\pawn_b.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "p"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num + 12
Next
'placing the black rocks (9, 16)
For Num = 9 To 16 Step 7
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\rock_b.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "r"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num - 8
Next
'placing the black horses (10, 15)
For Num = 10 To 15 Step 5
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\horse_b.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "h"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num - 8
Next
'placing the black officers (11, 14)
For Num = 11 To 14 Step 3
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\officer_b.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "o"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num - 8
Next
'placing the black queen (12)
Load FigImg(12)
FigImg(12).Picture = LoadPicture(ApPath & "pics\figures\queen_b.ico")
FigImg(12).Visible = True
FigImg(12).Enabled = False
FigImg(12).Tag = "q"
FigImg(12).ZOrder 0
PlaceFigAt FigImg(12).Index, 4
'placing the black king (13)
Load FigImg(13)
FigImg(13).Picture = LoadPicture(ApPath & "pics\figures\king_b.ico")
FigImg(13).Visible = True
FigImg(13).Enabled = False
FigImg(13).Tag = "k"
FigImg(13).ZOrder 0
PlaceFigAt FigImg(13).Index, 5

'placing the white pawns (17 - 24)
For Num = 17 To 24 Step 1
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\pawn_w.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "p"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num + 56
Next
'placing the white rocks (27, 32)
For Num = 25 To 32 Step 7
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\rock_w.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "r"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num + 60
Next
'placing the white horses (26, 31)
For Num = 26 To 31 Step 5
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\horse_w.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "h"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num + 60
Next
'placing the white officers (27, 30)
For Num = 27 To 30 Step 3
   Load FigImg(Num)
   FigImg(Num).Picture = LoadPicture(ApPath & "pics\figures\officer_w.ico")
   FigImg(Num).Visible = True
   FigImg(Num).Enabled = False
   FigImg(Num).Tag = "o"
   FigImg(Num).ZOrder 0
   PlaceFigAt FigImg(Num).Index, Num + 60
Next
'placing the white queen (28)
Load FigImg(28)
FigImg(28).Picture = LoadPicture(ApPath & "pics\figures\queen_w.ico")
FigImg(28).Visible = True
FigImg(28).Enabled = False
FigImg(28).Tag = "q"
FigImg(28).ZOrder 0
PlaceFigAt FigImg(28).Index, 61 + 27
'placing the white king (29)
Load FigImg(29)
FigImg(29).Picture = LoadPicture(ApPath & "pics\figures\king_w.ico")
FigImg(29).Visible = True
FigImg(29).Enabled = False
FigImg(29).Tag = "k"
FigImg(29).ZOrder 0
PlaceFigAt FigImg(29).Index, 59 + 30
'-----------------------
'placing the coordinates
'-----------------------
'placing the numbers 1 to 8
CrdLabel(0).Visible = False
For Num = 1 To 8 Step 1
    Load CrdLabel(Num)
    CrdLabel(Num).Caption = Num
    CrdLabel(Num).BackStyle = 0
    CrdLabel(Num).BorderStyle = 0
    CrdLabel(Num).Visible = True
    CrdLabel(Num).Enabled = True
    CrdLabel(Num).Height = 20
    CrdLabel(Num).Width = 12
    CrdLabel(Num).Top = Picture1.Top + (8 * 60) - (Num * 60) + 25
    CrdLabel(Num).Left = Picture1.Left - CrdLabel(Num).Width
Next

Dim Alph As String
Alph = "ABCDEFGH"
For Num = 1 To 8 Step 1
    Load CrdLabel(Num + 8)
    CrdLabel(Num + 8).Caption = Mid(Alph, Num, 1)
    CrdLabel(Num + 8).BackStyle = 0
    CrdLabel(Num + 8).BorderStyle = 0
    CrdLabel(Num + 8).Visible = True
    CrdLabel(Num + 8).Enabled = True
    CrdLabel(Num + 8).Height = 15
    CrdLabel(Num + 8).Width = 15
    CrdLabel(Num + 8).Top = Picture1.Top - CrdLabel(Num + 8).Height
    CrdLabel(Num + 8).Left = Picture1.Left + (Num * 60) - 35
Next
'---------------------------------------------
'creating the imageboxes for the moving colors
'---------------------------------------------
For Num = 1 To UBound(Box) Step 1
    Load MovBox(Num)
    MovBox(Num).Enabled = False
    MovBox(Num).Visible = True
    MovBox(Num).ZOrder 1
    MovBox(Num).Left = Box(Num).XStart
    MovBox(Num).Top = Box(Num).YStart
    MovBox(Num).Width = 60
    MovBox(Num).Height = 60
Next
'------------------------------------------------
'load the settings variables from the config file
'------------------------------------------------
Dim Buff As String
Dim TempStr() As String
If Dir(ApPath & "sett.txt", vbNormal) <> vbNullString Then
    Open ApPath & "sett.txt" For Input As #1
        Do Until EOF(1) = True
            Line Input #1, Buff
            TempStr = Split(Buff, "=", , vbTextCompare)
            CallByName Me, TempStr(0), VbLet, TempStr(1)
        Loop
    Close #1
End If
'----------------------------------------------------------
'updating the variables and controls with the loaded values
'----------------------------------------------------------
Shape1.BorderColor = SelColor

For Num = 1 To 96
    MovBox(Num).Visible = ShowMov
Next

Exit Sub
'------------------
'the error handling
'------------------
Err:
    MsgBox Err.Description, vbCritical, App.Title
    End

End Sub
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------
'save the settings variables to the config file
'----------------------------------------------
Open ApPath & "sett.txt" For Output As #1
    Print #1, "FigMovSp" & "=" & FigMovSp
    Print #1, "SelColor" & "=" & Shape1.BorderColor
    Print #1, "ShowMov" & "=" & ShowMov
    Print #1, "MyName" & "=" & MyName
    Print #1, "Snd" & "=" & Snd
    Print #1, "EndTurn" & "=" & EndTurn
Close
'------------
'close the ws
'------------
WS.CloseSck

'----------------
'unload all forms
'----------------
'Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6

End Sub
Private Sub menu11_Click(Index As Integer)
'start a new game
If MsgBox("Are you sure you want to stop this game?", vbQuestion & vbYesNo, App.Title) = vbYes Then
    StartNewGame
End If

End Sub
Private Sub menu12_Click(Index As Integer)
'Exit
Command2_Click

End Sub
Private Sub menu21_Click(Index As Integer)
'show the settings form
Form2.Show 1, Form1
    
End Sub
Public Sub menu31_Click(Index As Integer)
'show the about dialog
Form5.Show 1

End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'-----------------------------------------
'find which box number has been clicked at
'-----------------------------------------
Dim Num As Long
Num = 1
Do Until ((Box(Num).XStart <= x) And (x <= Box(Num).XStart + 60)) And ((Box(Num).YStart <= y) And (y <= Box(Num).YStart + 60))
    Num = Num + 1
Loop
'hide all movement colors
Dim BoxNum As Long
For BoxNum = 1 To UBound(Box) Step 1
    MovBox(BoxNum).Picture = LoadPicture()
Next

Select Case Button
Case 1:
    '---------------
    'select a figure
    '---------------
    'check if there's a figure in the box
    If Box(Num).ImgIndex > 0 Then
        'showing the selection box
        Shape1.Top = Box(Num).YStart
        Shape1.Left = Box(Num).XStart
        Shape1.Visible = True
        
        FromBoxIn = Num
        
        'show the moving colors for the selected figure
        ShowMovBox Num
        
    Else
        'deselect the figure if any is sellected
        Shape1.Visible = False
    End If
Case 2:
    '------------------------
    'move the selected figure
    '------------------------
    'show the popup menu of changing the figure type
    If Shape1.Visible = False And Box(Num).ImgIndex > 0 And Box(Num).IsSafe = False Then
        Form3.FigT = FigImg(Box(Num).ImgIndex).Tag
        Form3.FigC = FindFigColor(Box(Num).ImgIndex)
        Form3.FigInt = Box(Num).ImgIndex
        PopupMenu popup1
    End If
    'check if we're trying to hit a figure:
    ' - from the safe to the safe, or
    ' - from the safe to the board, or
    ' - from the board to the safe
    If _
        Box(FromBoxIn).IsSafe = True _
        And Box(Num).IsSafe = True _
        And Box(Num).ImgIndex > 0 _
    Or _
        Box(FromBoxIn).IsSafe = True _
        And Box(Num).IsSafe = False _
        And Box(Num).ImgIndex > 0 _
    Or _
        Box(FromBoxIn).IsSafe = False _
        And Box(Num).IsSafe = True _
        And Box(Num).ImgIndex > 0 _
    Then
        GoTo 1
    End If
    'make sure we don't put rival figures in our safe
    If _
        (FindFigColor(Box(FromBoxIn).ImgIndex) = "b" And Num >= 57 And Box(Num).IsSafe = True) _
    Or _
        (FindFigColor(Box(FromBoxIn).ImgIndex) = "w" And Num < 57 And Box(Num).IsSafe = True) _
    Then
        GoTo 1
    End If
    'check if there's a figure in the destination box and is something selected
    If Shape1.Visible = True And Box(FromBoxIn).ImgIndex > 0 Then
        'check if there's a figure from the same color in the box
        If FindFigColor(Box(Num).ImgIndex) = FindFigColor(Box(FromBoxIn).ImgIndex) Then
            'the figure is from the same color - hide the selection box
            GoTo 1
        Else
            'send the data to the other player
            If IsOnline = True Then
                WS.SendData "mov:" & Box(FromBoxIn).ImgIndex & ":" & FromBoxIn & ":" & Num
                'disallow command5 until the figure is moving
                Command5.Enabled = False
            End If
            Picture1.Enabled = False
            'move the figure
            MoveFigTo Box(FromBoxIn).ImgIndex, FromBoxIn, Num
            'hide the selection box
            Shape1.Visible = False
        End If
    Else
1:      'nothing in the box - hide the selection box
        Shape1.Visible = False
    End If
End Select

End Sub
Private Sub popup11_Click()
'show the choose figure menu
Form3.Show 1, Form1

End Sub
Private Sub Text1_Change()
'show the end of the text
Text1.SelStart = Len(Text1.Text)

End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'use enter as "send"
If KeyCode = vbKeyReturn Then
    Command1_Click
End If

End Sub
Private Sub Timer1_Timer()
DoEvents
'check which speed we will use
If IsOnline = True Then
    MovSp = OnlineFigMovSp
Else
    MovSp = FigMovSp
End If
'check if the figure reached the destination
If (FigImg(FigIn).Left = Box(ToBox).XStart) And (FigImg(FigIn).Top = Box(ToBox).YStart) Then
    Timer1.Enabled = False
    'reenable picture1
    If IsOnline = False Then
        Picture1.Enabled = True 'we're offline
    Else    'we're online
        'is it my turn?
        If MyTurn = True Then 'yep
            Picture1.Enabled = True
            Command5.Enabled = True
        End If
    End If
Else
    'check if the X is descending or ascending
    Select Case FigImg(FigIn).Left
    Case Is < Box(ToBox).XStart:
        FigImg(FigIn).Left = FigImg(FigIn).Left + MovSp
    Case Is > Box(ToBox).XStart:
        FigImg(FigIn).Left = FigImg(FigIn).Left - MovSp
    End Select
    'check if the Y is descending or ascending
    Select Case FigImg(FigIn).Top
    Case Is < Box(ToBox).YStart:
        FigImg(FigIn).Top = FigImg(FigIn).Top + MovSp
    Case Is > Box(ToBox).YStart:
        FigImg(FigIn).Top = FigImg(FigIn).Top - MovSp
    End Select
End If

End Sub
Private Sub PlaceFigAt(ImgInt As Long, BoxNum As Long)
DoEvents
'place a figure to a desired box
FigImg(ImgInt).Left = Box(BoxNum).XStart
FigImg(ImgInt).Top = Box(BoxNum).YStart
Box(BoxNum).ImgIndex = FigImg(ImgInt).Index

End Sub
Private Sub MoveFigTo(ImgInt As Long, FromBoxInt As Long, ToBoxInt As Long)
DoEvents
'move the figure to the second box
FigIn = ImgInt
ToBox = ToBoxInt
Timer1.Enabled = True

'place the figure if any, from the "to" box to the safe
If Box(ToBoxInt).ImgIndex > 0 Then
    PlaceInSafe Box(ToBoxInt).ImgIndex
End If

'set the new information in the "To" box
Box(ToBoxInt).ImgIndex = FigImg(ImgInt).Index

'remove the information from the "from" box
Box(FromBoxInt).ImgIndex = 0

End Sub
Private Function FindFigColor(ImgInt As Long) As String
DoEvents
'find the color of the selected figure
If (FigImg(ImgInt).Index <= 16) And (FigImg(ImgInt).Index > 0) Then
    FindFigColor = "b"  'black color
Else
    FindFigColor = "n"  'no color
End If
If FigImg(ImgInt).Index >= 17 Then
    FindFigColor = "w"  'white color
End If

End Function
Private Sub PlaceInSafe(ImgInt As Long)
DoEvents
Dim Num As Long
Dim ToNum
If FindFigColor(ImgInt) = "b" Then
    Num = 1
    ToNum = 16
Else
    Num = 17
    ToNum = 32
End If
'check every safe box for a free space
Do Until Num > ToNum
    If Box(SBox(Num)).ImgIndex <= 0 Then
    'the box is free
        PlaceFigAt ImgInt, SBox(Num)
        Exit Sub
    End If
    Num = Num + 1
Loop

End Sub
Private Sub ShowMovBox(BoxInt As Long)
'------------------------------
'display the allowed move zones
'------------------------------
Dim L As Long
Dim R As Long
Dim U As Long
Dim D As Long
Dim LU As Long
Dim LD As Long
Dim RU As Long
Dim RD As Long
Dim Num As Long
'find the row on which the figure is located
For Num = 0 To 7
    If BoxInt > Num * 12 And BoxInt <= (Num * 12) + 12 Then
        L = BoxInt - (Num * 12) - 1 'L
        R = (Num + 1) * 12 - BoxInt 'R
        U = Num                     'U
        D = 8 - (Num + 1)           'D
        If U <= L Then              'LU
            LU = U
        Else
            LU = L
        End If
        If D <= L Then              'LD
            LD = D
        Else
            LD = L
        End If
        If U <= R Then              'RU
            RU = U
        Else
            RU = R
        End If
        If D <= R Then              'RD
            RD = D
        Else
            RD = R
        End If
    End If
Next
'check if the figure is in the safe
If Box(BoxInt).IsSafe = False Then
    Select Case FigImg(Box(BoxInt).ImgIndex).Tag
    Case "k":   'king
        If L > 0 Then L = 1
        GoSub L
        If R > 0 Then R = 1
        GoSub R
        If U > 0 Then U = 1
        GoSub U
        If D > 0 Then D = 1
        GoSub D
        If LU > 0 Then LU = 1
        GoSub LU
        If LD > 0 Then LD = 1
        GoSub LD
        If RU > 0 Then RU = 1
        GoSub RU
        If RD > 0 Then RD = 1
        GoSub RD
    Case "q":   'queen
        GoSub L
        GoSub R
        GoSub U
        GoSub D
        GoSub LU
        GoSub LD
        GoSub RU
        GoSub RD
    Case "o":   'officer
        GoSub LU
        GoSub LD
        GoSub RU
        GoSub RD
    Case "h":   'horse
        GoSub H
    Case "r":   'rock
        GoSub L
        GoSub R
        GoSub U
        GoSub D
    Case "p":   'pawn
        GoSub P
    End Select
End If

Exit Sub
'----------------------------------------------------------------------------------
L:  'left
For Num = 1 To L Step 1
    If Box(BoxInt - Num).ImgIndex <= 0 Then 'check is there a figure on the box
    'no figure there
        MovBox(BoxInt - Num).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt - Num).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt - Num).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
R:  'right
For Num = 1 To R Step 1
    If Box(BoxInt + Num).ImgIndex <= 0 Then 'check is there a figure on the box
    'no figure there
        MovBox(BoxInt + Num).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt + Num).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt + Num).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
U:  'up
For Num = 1 To U Step 1
    'check is there a figure on the box
    If Box(BoxInt - (Num * 12)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt - (Num * 12)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt - (Num * 12)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt - (Num * 12)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
D:  'down
For Num = 1 To D Step 1
    'check is there a figure on the box
    If Box(BoxInt + (Num * 12)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt + (Num * 12)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt + (Num * 12)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt + (Num * 12)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
LU:  'left-up
For Num = 1 To LU Step 1
    'check is there a figure on the box
    If Box(BoxInt - (Num * 13)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt - (Num * 13)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt - (Num * 13)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt - (Num * 13)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
LD:  'left-down
For Num = 1 To LD Step 1
    'check is there a figure on the box
    If Box(BoxInt + (Num * 11)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt + (Num * 11)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt + (Num * 11)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt + (Num * 11)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
RU:  'right-up
For Num = 1 To RU Step 1
    'check is there a figure on the box
    If Box(BoxInt - (Num * 11)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt - (Num * 11)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt - (Num * 11)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt - (Num * 11)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
RD:  'right-down
For Num = 1 To RD Step 1
    'check is there a figure on the box
    If Box(BoxInt + (Num * 13)).ImgIndex <= 0 Then
    'no figure there
        MovBox(BoxInt + (Num * 13)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
    Else
    'yep, there is a figure
        'check if the figure is a rival one
        If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt + (Num * 13)).ImgIndex) Then
        'the figure is a rival one, display a red movebox
            MovBox(BoxInt + (Num * 13)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            Return
        Else
        'the figure is friendly, no box displayed
            Return
        End If
    End If
Next
Return
'----------------------------------------------------------------------------------
H:  'a special section for the horse's movements
Dim Add(1 To 8) As Long
Add(1) = -14
Add(2) = -25
Add(3) = -23
Add(4) = -10
Add(5) = 14
Add(6) = 25
Add(7) = 23
Add(8) = 10

For Num = 1 To 8 Step 1
    'check if we're not exceeding the bounds
    If BoxInt + Add(Num) > LBound(Box) And BoxInt + Add(Num) < UBound(Box) Then
        'check is there a figure on the box
        If Box(BoxInt + Add(Num)).ImgIndex <= 0 Then
        'no figure there
            MovBox(BoxInt + Add(Num)).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
        Else
        'yep, there is a figure
            'check if the figure is a rival one
            If FindFigColor(Box(BoxInt).ImgIndex) <> FindFigColor(Box(BoxInt + Add(Num)).ImgIndex) Then
            'the figure is a rival one, display a red movebox
                MovBox(BoxInt + Add(Num)).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
            End If
        End If
    End If
Next
Return
'----------------------------------------------------------------------------------
P:  'a special section for the pawn's movements
'determine the pawn's color
If FindFigColor(Box(BoxInt).ImgIndex) = "b" Then 'black
    'check if we're not exceeding the bounds (for movement)
    If BoxInt + 12 <= UBound(Box) Then
        'check if there's a figure in front of the pawn
        If Box(BoxInt + 12).ImgIndex <= 0 Then
            MovBox(BoxInt + 12).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
        End If
    End If
    'check if we're not exceeding the bounds (for attack #1)
    If BoxInt + 11 <= UBound(Box) Then
        'determine the color of the pawn we're trying to attack
        If FindFigColor(Box(BoxInt + 11).ImgIndex) = "w" Then
            MovBox(BoxInt + 11).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
        End If
    End If
    'check if we're not exceeding the bounds (for attack #2)
    If BoxInt + 13 <= UBound(Box) Then
        'determine the color of the pawn we're trying to attack
        If FindFigColor(Box(BoxInt + 13).ImgIndex) = "w" Then
            MovBox(BoxInt + 13).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
        End If
    End If
Else    'white
    'check if we're not exceeding the bounds (for movement)
    If BoxInt - 12 >= LBound(Box) Then
        'check if there's a figure in front of the pawn
        If Box(BoxInt - 12).ImgIndex <= 0 Then
            MovBox(BoxInt - 12).Picture = LoadPicture(ApPath & "pics\tr_green.ico")
        End If
    End If
    'check if we're not exceeding the bounds (for attack #1)
    If BoxInt - 11 >= LBound(Box) Then
        'determine the color of the pawn we're trying to attack
        If FindFigColor(Box(BoxInt - 11).ImgIndex) = "b" Then
            MovBox(BoxInt - 11).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
        End If
    End If
    'check if we're not exceeding the bounds (for attack #2)
    If BoxInt - 13 >= LBound(Box) Then
        'determine the color of the pawn we're trying to attack
        If FindFigColor(Box(BoxInt - 13).ImgIndex) = "b" Then
            MovBox(BoxInt - 13).Picture = LoadPicture(ApPath & "pics\tr_red.ico")
        End If
    End If
End If
Return
'----------------------------------------------------------------------------------

End Sub
Private Sub WS_CloseSck()
'show error message
MsgBox "The connection to the recipient was terminated.", vbCritical, App.Title
RtnMain

End Sub
Private Sub WS_Connect()
'----------------------------------
'hide form4 and initialize the game
'----------------------------------
IsOnline = True

'hide the main menu
Unload Form4
'show the lobby
Form6.Show , Form1
'show my name in the lobby
Form6.Label2.Caption = MyName
'send my name to the host
WS.SendData "mna:" & MyName

End Sub
Private Sub WS_ConnectionRequest(ByVal requestID As Long)
'--------------------------------------------------------------
'accept the incoming connection, if we're not currently playing
'--------------------------------------------------------------
If IsOnline = False Then
    WS.CloseSck
    WS.Accept requestID
    IsOnline = True
    
    'hide the main menu
    Unload Form4
    'show the lobby
    Form6.Show , Form1
    'show my name in the lobby
    Form6.Label1.Caption = MyName
    'send my name to the host
    WS.SendData "mna:" & MyName

End If

End Sub
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
'-------------------------
'proceed the incoming data
'-------------------------
DoEvents
WS.GetData Dat
Dim TempStr() As String

1: Select Case Mid(Dat, 1, 3)
    Case "mov":
        'temporary disable picture1 until the moving is done
        Picture1.Enabled = False
        'do the moving
        TempStr = Split(Dat, ":", , vbTextCompare)
        MoveFigTo CLng(TempStr(1)), CLng(TempStr(2)), CLng(TempStr(3))
        
    Case "fig":
        'change the figure's image and type
        TempStr = Split(Dat, ":", , vbTextCompare)
        FigImg(TempStr(1)).Tag = TempStr(3)
        Select Case TempStr(3)
        Case "p":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\pawn_" & TempStr(2) & ".ico")
        Case "k":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\king_" & TempStr(2) & ".ico")
        Case "q":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\queen_" & TempStr(2) & ".ico")
        Case "o":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\officer_" & TempStr(2) & ".ico")
        Case "h":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\horse_" & TempStr(2) & ".ico")
        Case "r":
            FigImg(TempStr(1)).Picture = LoadPicture(ApPath & "pics\figures\rock_" & TempStr(2) & ".ico")
        End Select
        'show notification in the chat panel
        Text1.Text = Text1.Text & vbNewLine & "-- " & TempStr(4) & " has changed a figure."
        'play a notification sound
        If Snd = True Then
            PlaySound ApPath & "sound\fig.wav", vbEmpty, SND_FILENAME & SND_ASYNC
        End If
        
    Case "msg":
        'add the message to the textbox
        Text1.Text = Text1.Text & vbNewLine & Mid(Dat, 5)
        
    Case "chk":
        TempStr = Split(Dat, ":", , vbTextCompare)
        'play the check sound
        If Snd = True Then
            PlaySound ApPath & "sound\check.wav", vbEmpty, SND_FILENAME & SND_ASYNC
        End If
        Text1.Text = Text1.Text & vbNewLine & "-- " & TempStr(1) & " says CHECK!"
        
    Case "csl":
        TempStr = Split(Dat, ":", , vbTextCompare)
        'play the castling sound
        If Snd = True Then
            PlaySound ApPath & "sound\castling.wav", vbEmpty, SND_FILENAME & SND_ASYNC
        End If
        Text1.Text = Text1.Text & vbNewLine & "-- " & TempStr(1) & " says CASTLING!"
        
    Case "chm":
        TempStr = Split(Dat, ":", , vbTextCompare)
        'play the checkmate sound
        If Snd = True Then
            PlaySound ApPath & "sound\checkmate.wav", vbEmpty, SND_FILENAME & SND_ASYNC
        End If
        Text1.Text = Text1.Text & vbNewLine & "-- " & TempStr(1) & " says CHECKMATE!"
        
        
        
    Case "mna":
        TempStr = Split(Dat, ":", , vbTextCompare)
        'show the other player's name in the lobby
        If IsServer = True Then
            Form6.Label2.Caption = TempStr(1)
        Else
            Form6.Label1.Caption = TempStr(1)
        End If
        YourName = TempStr(1)
        
    Case "clr":
        TempStr = Split(Dat, ":", , vbTextCompare)
        'show the other player's color in the lobby
        If IsServer = True Then
            Form6.Label4.BackColor = TempStr(1)
        Else
            Form6.Label3.BackColor = TempStr(1)
        End If
        
    Case "lms"::
        'receive a message in the lobby
        Form6.Text1.Text = Form6.Text1.Text & vbNewLine & Mid(Dat, 5)
    
    Case "sta":
        'show the board
        Unload Form6
        Me.Show
        'start the game
        StartMulGame
        
    Case "end":
        'start my turn
        MyTurn = True
        
        'enable my controls
        Picture1.Enabled = True
        Command5.Enabled = True
        
        'enable the notifications
        Command3.Enabled = True
        Command4.Enabled = True
        Command6.Enabled = True
        
        'update the color in the colorbox
        If MyColor = "w" Then
            Label4.ForeColor = &HFFFFFF
        Else
            Label4.ForeColor = &H0&
        End If
        Label4.Caption = MyName
        'show message
        If EndTurn = True Then
            MsgBox "It's your turn now.", vbInformation, App.Title
        End If
        'play the sound
        If Snd = True Then
            PlaySound ApPath & "sound\turn.wav", vbEmpty, SND_FILENAME & SND_ASYNC
        End If

End Select

End Sub
Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'display an error message
MsgBox "WinSock error:" & vbNewLine & Number & " : " & Description, vbCritical, App.Title
RtnMain

End Sub
Private Sub StartNewGame()
'--------------------------------
'stop all games currently running
'--------------------------------
RtnMain

End Sub
Public Sub RtnMain()
'close the socket
WS.CloseSck
IsOnline = False
'unload all forms
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form1
'show main menu
Form4.Show 1

End Sub
Public Sub StartMulGame()
'start the multiplayer game
If MyColor = "w" Then
    Label4.ForeColor = &HFFFFFF
    Label4.Caption = MyName
    
    Picture1.Enabled = True
    Command5.Enabled = True
    
    'enable the notifications
    Command3.Enabled = True
    Command4.Enabled = True
    Command6.Enabled = True

    'show message
    If EndTurn = True Then
        MsgBox "It's your turn now.", vbInformation, App.Title
    End If
    
    'play the sound
    If Snd = True Then
        PlaySound ApPath & "sound\turn.wav", vbEmpty, SND_FILENAME & SND_ASYNC
    End If
    
    MyTurn = True
Else
    Label4.ForeColor = &HFFFFFF
    Label4.Caption = YourName
    
    'disable the notifications
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    
    Picture1.Enabled = False
    Command5.Enabled = False
    
    MyTurn = False
End If

End Sub
