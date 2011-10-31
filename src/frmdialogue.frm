VERSION 5.00
Begin VB.Form frmdialogue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialogue"
   ClientHeight    =   4920
   ClientLeft      =   1110
   ClientTop       =   1395
   ClientWidth     =   1455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      ToolTipText     =   "This is where you can take notes on what is happening in the adventure"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   4560
   End
   Begin VB.CommandButton Cmdtalk 
      Caption         =   "Talk!"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "This button will make your character say what is selected above"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Click one of these options to choose what to say"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Click one of these options to choose what to say"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Click one of these options to choose what to say"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Click one of these options to choose what to say"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdnotes 
      Caption         =   "Notes"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Click on this button to go to the notes section."
      Top             =   4530
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "This is where all of the dialogue in the game appears"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblnotes 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Speech Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lbldialogue 
      Caption         =   "Dialogue:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmdialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MovedLeft As Boolean

Private Sub cmdprevtext_Click()

'If dialoguebuttonstate = False Then
'    curtext = Text1.Text
'    Text1.Text = prevtext
'    cmdprevtext.Caption = "Current Text"
'    dialoguebuttonstate = True
'Else
'    Text1.Text = curtext
'    cmdprevtext.Caption = "Previous Text"
'    dialoguebuttonstate = False
'End If

End Sub

Private Sub cmdnotes_Click()
Timer1.Enabled = True
Slidesoundstart
End Sub

Private Sub Form_Load()
Randomize
'frmdialogue.Height = frmimage.Height
Speechoptionsdisabled
MovedLeft = False

End Sub

Private Sub Movecontrolsleft()
Text1.Move Text1.Left - 50
lbldialogue.Move lbldialogue.Left - 50
Label1.Move Label1.Left - 50
Option1.Move Option1.Left - 50
Option2.Move Option2.Left - 50
Option3.Move Option3.Left - 50
Option4.Move Option4.Left - 50
Cmdtalk.Move Cmdtalk.Left - 50
Text2.Move Text2.Left - 50
lblnotes.Move lblnotes.Left - 50
If Text2.Left < 50 Then
    Timer1.Enabled = False
    cmdnotes.Caption = "Dialogue"
    MovedLeft = True
    Moveleft
End If

End Sub
Private Sub MoveControlsLeftBack()

Text1.Move Text1.Left + 50
lbldialogue.Move lbldialogue.Left + 50
Label1.Move Label1.Left + 50
Option1.Move Option1.Left + 50
Option2.Move Option2.Left + 50
Option3.Move Option3.Left + 50
Option4.Move Option4.Left + 50
Cmdtalk.Move Cmdtalk.Left + 50
Text2.Move Text2.Left + 50
lblnotes.Move lblnotes.Left + 50

If Text2.Left > 1550 Then
    Timer1.Enabled = False
    cmdnotes.Caption = "Notes"
    MovedLeft = False
    Moveleft
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frminv.statbar1.SimpleText = ""
End Sub

Private Sub Text1_Change()
If MovedLeft = True Then
Timer1.Enabled = True
Slidesoundstart
End If
End Sub

Private Sub Timer1_Timer()
If MovedLeft = False Then
Movecontrolsleft
ElseIf MovedLeft = True Then
MoveControlsLeftBack
End If

End Sub
