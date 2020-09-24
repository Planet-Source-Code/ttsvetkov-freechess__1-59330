VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3840
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'--
'OK
'--
Dim Num As Long
'save the figure moving speed
Form1.FigMovSp = Combo1.Text

'save the selection box color
Form1.Shape1.BorderColor = Shape1.BorderColor

'show/hide the movement colors
If Check1.Value = 1 Then
    Form1.ShowMov = True
Else
    Form1.ShowMov = False
End If
For Num = 1 To 96
    Form1.MovBox(Num).Visible = Form1.ShowMov
Next

'save my name
Form1.MyName = Text1.Text

'play sound swich
If Check2.Value = 1 Then
    Form1.Snd = True
Else
    Form1.Snd = False
End If

'show message swich
If Check3.Value = 1 Then
    Form1.EndTurn = True
Else
    Form1.EndTurn = False
End If

Unload Me

End Sub
Private Sub Command2_Click()
'------
'Cancel
'------
Unload Me

End Sub
Private Sub Command3_Click()
'--------------
'choose a color
'--------------
CD1.ShowColor
Shape1.BorderColor = CD1.Color

End Sub
Private Sub Form_Load()
'--------------------
'setting the controls
'--------------------
Me.Caption = "Settings"
Me.ScaleMode = 3
Command1.Caption = "OK"
Command2.Caption = "Cancel"

'setting the figspeed options
Label1.Caption = "The moving speed is valid only for Singleplay. In Multiplayer the speed is allways 10, making sure the moves are synchronized with the both players."
Frame1.Caption = "Figure moving speed"
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "10"
Combo1.AddItem "20"
Combo1.AddItem "30"
Combo1.AddItem "60"
Combo1.Text = Form1.FigMovSp

'setting the choose color options
Frame2.Caption = "Selection box color"
Command3.Caption = "Choose..."
Shape1.BorderColor = Form1.Shape1.BorderColor

'setting the show movement colors
If Form1.ShowMov = True Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
Check1.Caption = "Display movement colors"

'setting the port number
Frame3.Caption = "My Name"
Text1.Text = Form1.MyName

'play sound
Check2.Caption = "Play Sounds"
If Form1.Snd = True Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If

'end turn
Check3.Caption = "Show ""End Turn"" messages"
If Form1.EndTurn = True Then
    Check3.Value = 1
Else
    Check3.Value = 0
End If

End Sub
