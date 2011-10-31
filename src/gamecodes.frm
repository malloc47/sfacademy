VERSION 5.00
Begin VB.Form gamecodes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Close Window After Load"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtname 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "gamecodes.frx":0000
      Left            =   1320
      List            =   "gamecodes.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lbllevelnam 
      Height          =   1815
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Game Codes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1493
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "gamecodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CheckTheCheck()
If Check1.Value = 1 Then gamecodes.Visible = False
End Sub

Private Sub cmdcancel_Click()
gamecodes.Visible = False
End Sub

Private Sub Form_Activate()

If NUMCodes = 1 Then
gamecodes.List1.Clear
gamecodes.List1.AddItem ("GB457HTI")
ElseIf NUMCodes = 2 Then
gamecodes.List1.Clear
gamecodes.List1.AddItem ("GB457HTI")
gamecodes.List1.AddItem ("YH243IKT")
End If

End Sub

Private Sub List1_Click()
txtname.Text = List1.Text
End Sub

Private Sub Load_Click()
If txtname.Text = "GB457HTI" Then
    Screen1 = 1
    frmimage.Image1.Picture = LoadPicture(App.Path & "\images\hallmain1.jpg")
    Door1Open = 0
    ReSound = False
    Elevator1Open = 0
    PipeTaken = False
    PipeAV = False
    frmimage.imgfront.Enabled = True
    frmimage.Small
    frminv.imgpipe.Visible = False
    CheckTheCheck
ElseIf txtname.Text = "YH243IKT" Then
    If NUMCodes = 1 Then
    gamecodes.List1.AddItem ("YH243IKT")
    NUMCodes = 2
    End If
    
    Screen1 = 21
    frmimage.Image1.Picture = LoadPicture(App.Path & "\images\halltop1.jpg")
    ElevatorDoorsOpen = True
    PipeTaken = True
    PipeAV = True
    frminv.imgpipe.Visible = True
    frmimage.imgfront.Enabled = True
    frmimage.Large
    Button2Clicked = True
    CheckTheCheck
Else

End If

End Sub

Private Sub txtname_Change()
If txtname.Text = "GB457HTI" Then
lbllevelnam.Caption = "Beginning of game"
ElseIf txtname.Text = "YH243IKT" Then
lbllevelnam.Caption = "Top of elevator, above welcome center"
Else
lbllevelnam.Caption = ""
End If

End Sub
