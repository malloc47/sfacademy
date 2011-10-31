VERSION 5.00
Begin VB.Form frmimage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starfleet Academy"
   ClientHeight    =   3600
   ClientLeft      =   2655
   ClientTop       =   1680
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox ElevatorLight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1890
      Picture         =   "main1.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer ElevatorControl 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   960
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   360
      Top             =   3120
   End
   Begin VB.Image imgfront 
      Height          =   1455
      Left            =   1680
      MouseIcon       =   "main1.frx":0426
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image debris 
      Height          =   855
      Left            =   240
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgright 
      Height          =   3615
      Left            =   4560
      MouseIcon       =   "main1.frx":0730
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgleft 
      Height          =   3615
      Left            =   0
      MouseIcon       =   "main1.frx":0A3A
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   0
      Picture         =   "main1.frx":0D44
      Top             =   0
      Width           =   4800
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuoptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusaved 
         Caption         =   "&Saved Positions"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusettingsoptions 
         Caption         =   "Settings"
         Begin VB.Menu mnuwindowsave 
            Caption         =   "S&ave Settings"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnurestoresaved 
            Caption         =   "Restore Settings"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnusound 
      Caption         =   "Sound"
      Begin VB.Menu mnumusic 
         Caption         =   "&Music"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnusoundon 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnucdmusic 
         Caption         =   "CD Music"
      End
      Begin VB.Menu mnuchoosesong 
         Caption         =   "&Choose Song..."
         Begin VB.Menu mnusti 
            Caption         =   "Star Trek I - The Motion Picture"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnustii 
            Caption         =   "Star Trek II - The Wrath of Khan"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnustiii 
            Caption         =   "Star Trek III - The Search for Spock"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnustiv 
            Caption         =   "Star Trek IV - The Journey Home"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnusttng 
            Caption         =   "Star Trek - The Next Generation"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnustvi 
            Caption         =   "Star Trek VI - The Undiscovered Country"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnustvii 
            Caption         =   "Star Trek VII - Generations"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnustviii 
            Caption         =   "Star Trek VIII - First Contact"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnustds9 
            Caption         =   "Star Trek - Deep Space 9"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnustvoy 
            Caption         =   "Star Trek - Voyager"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnusttos 
            Caption         =   "Star Trek - Origional Series"
            Shortcut        =   {F12}
         End
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
      Begin VB.Menu mnurestore 
         Caption         =   "&Restore Window Positions"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub WelcomeDialogue()
If DialogueToggle = False Then
frmdialogue.Text1.Text = "This place is a mess, I guess they were attacked all right."
DialogueToggle = True
End If
End Sub
Public Sub Bottomalign()
    imgfront.Height = 375
    imgfront.Left = 240
    imgfront.Top = 2880
    imgfront.Width = 4335
End Sub
Public Sub Large()
    imgfront.Height = 3375
    imgfront.Left = 360
    imgfront.Top = 120
    imgfront.Width = 4095
End Sub
Public Sub Small()
    imgfront.Height = 1455
    imgfront.Left = 1680
    imgfront.Top = 1080
    imgfront.Width = 1455
End Sub
Private Sub Elevatorpic()

If ButtonClicked = True Then
    If Elevator1Open = 0 Then
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain5cl.jpg")
        ElseIf Elevator1Open = 1 Then
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain5jammed.jpg")
        ElseIf Elevator1Open = 2 Then
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain5.jpg")
    End If
ElseIf ButtonClicked = False Then
Image1.Picture = LoadPicture(App.Path & "\images\hallmain5cl.jpg")

End If
End Sub

Private Sub Pipescreen()
If PipeTaken = False Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1c.jpg")
    imgfront.Enabled = True
ElseIf PipeTaken = True Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1d.jpg")
    imgfront.Enabled = False
End If

End Sub
Public Sub Pipesize()
    imgfront.Height = 615
    imgfront.Left = 240
    imgfront.Top = 2280
    imgfront.Width = 4335
End Sub
Public Sub Buttonsize()
imgfront.Height = 300
imgfront.Left = 2570
imgfront.Top = 1100
imgfront.Width = 255
End Sub
Public Sub Buttonsize2()
imgfront.Height = 255
imgfront.Left = 1350
imgfront.Top = 1260
imgfront.Width = 255
End Sub
Private Sub Elevatorinsidepic()
If ElevatorDoorsOpen = False Then
    If Button2Clicked = False Then
        Image1.Picture = LoadPicture(App.Path & "\images\elevator1.jpg")
    Else
        Image1.Picture = LoadPicture(App.Path & "\images\elevator1a.jpg")
    End If
Else
    Image1.Picture = LoadPicture(App.Path & "\images\halltop1.jpg")
    imgfront.Enabled = True
End If

End Sub

Private Sub debris_DragDrop(Source As Control, x As Single, y As Single)
If Source.Tag = "eye" And Screen1 = 8 Then
WelcomeDialogue
MessageBox = MsgBox("This is debris.  It must have fallen from the celing.  I guess the distress calls from here were true, they were attacked.", vbInformation, "Pile of Debris")
Source.DragIcon = frminv.eye1.DragIcon
End If

End Sub

Private Sub debris_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If Source.Tag = "eye" Then
If Screen1 = 8 Then
    If State = 0 Then
        Source.DragIcon = frminv.eye2.Picture
    ElseIf State = 1 Then
        Source.DragIcon = frminv.eye1.DragIcon
    ElseIf State = 2 Then
        Source.DragIcon = frminv.eye2.Picture
    End If
Else
Source.DragIcon = frminv.eye1.DragIcon
End If
End If
End Sub

Private Sub debris_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Screen1 = 8 Then
frminv.statbar1.SimpleText = "A Pile of Debris"
Else
frminv.statbar1.SimpleText = ""
End If

End Sub

Private Sub ElevatorControl_Timer()
If ElevatorControl.Interval = 1200 Then
    WavFile = "\sound\elevato2.wav"
    PlaySoundLoop
    ReSound = True
    ElevatorLight.Visible = True
    ElevatorControl.Interval = 50

ElseIf ElevatorControl.Interval = 500 Then
    WavFile = "\sound\unjam.wav"
    PlaySound
    ReSound = False
    ElevatorControl.Enabled = False
Else
    If Screen1 = 21 Then
    ElevatorLight.Visible = False
    Else
    ElevatorLight.Visible = True
    End If

    ElevatorControl.Interval = 50
    ElevatorLight.Top = ElevatorLight.Top + 100
    If ElevatorLight.Top > 2365 Then
        ElevatorControl.Interval = 600
        ElevatorLight.Visible = False
        ElevatorLight.Top = 1200
    End If
    Elevator1Count = Elevator1Count + 1
    If Elevator1Count = 150 Then
    WavFile = "\sound\elevatostop.wav"
    PlaySound
    ElevatorControl.Interval = 500
    If Screen1 = 21 Then
        ElevatorLight.Visible = False
        Image1.Picture = LoadPicture(App.Path & "\images\halltop1.jpg")
    End If
    ElevatorDoorsOpen = True
    NUMCodes = 2
    End If

End If
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If Source.Tag = "pipe" Then Source.DragIcon = frminv.imgpipe2.DragIcon
End Sub

Private Sub Form_Load()
'Left = (Screen.Width - Width) \ 2
'Top = (Screen.Height - Height) \ 2
Load frmdialogue
Load frminv
'frmdialogue.Text1.Text = frmimage.Left & "Left" & frmimage.Top & "Left"
INIGet
Load gamecodes
gamecodes.Visible = False
MusicNum = 1
SongNum = 1
Screen1 = 1
Door1Open = 0
DialogueToggle = False
If mnumusic.Checked = True Then
    WaveCheck = waveOutGetNumDevs()
    If WaveCheck > 0 Then
        MusicOn = True
        Songchange
'        SoundOn = True
    Else
        MusicOn = False
        SoundOn = False
        mnusound.Enabled = False
'        mnusoundon.Checked = False
        mnumusic.Checked = False
    End If
End If
ReSound = False
Elevator1Open = 0
PipeTaken = False
PipeAV = False
Image2.Picture = LoadPicture(App.Path & "\images\hallmain5pipe.jpg")
frmdialogue.Visible = True
frminv.Visible = True
Songchange
End Sub

Private Sub Form_Resize()
If frmimage.WindowState = 1 Then
frminv.Visible = False
frmdialogue.Visible = False
frminv.WindowState = 1
frmdialogue.WindowState = 1


ElseIf frmimage.WindowState = 0 Then
frmdialogue.WindowState = 0
frminv.WindowState = 0
frmdialogue.Visible = True
frminv.Visible = True

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
INIMake
Unload frmdialogue
frmdialogue.Visible = False
Unload frminv
frminv.Visible = False
CloseMidi
WavFile = ""
PlaySoundLoop
End

End Sub
Private Sub Image1_DragDrop(Source As Control, x As Single, y As Single)
If Source.Tag = "pipe" Then Source.DragIcon = frminv.imgpipe2.DragIcon
If Source.Tag = "eye" Then Source.DragIcon = frminv.eye1.DragIcon

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frminv.statbar1.SimpleText = ""
End Sub

Private Sub imgfront_DragDrop(Source As Control, x As Single, y As Single)
If Screen1 = 10 Then
    If Elevator1Open = 1 Then
        If Source.Tag = "pipe" Then
            Elevator1Open = 2
            Elevatorpic
            PipeUsed = True
            Playsoundunjam
            Source.DragIcon = frminv.imgpipe2.DragIcon
        ElseIf Source.Tag = "eye" Then
            MessageBox = MsgBox("This is the door to the elevator.  It appears to be jammed, probobly from the attack.  If I could just find something to pry it open...", vbInformation, "A Jammed Door")
        End If
    End If
Else
If Source.Tag = "pipe" Then Source.DragIcon = frminv.imgpipe2.DragIcon
End If

End Sub

Private Sub imgfront_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If Source.Tag = "pipe" Then
If Screen1 = 10 And Elevator1Open = 1 Then
    If State = 0 Then
        Source.DragIcon = frminv.imgpipe.Picture
        Image1.Picture = Image2.Picture
    ElseIf State = 1 Then
        Source.DragIcon = frminv.imgpipe2.DragIcon
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain5jammed.jpg")
    ElseIf State = 2 Then
        Source.DragIcon = frminv.imgpipe.Picture
        Image1.Picture = Image2.Picture
    End If
Else
Source.DragIcon = frminv.imgpipe2.DragIcon
End If

ElseIf Source.Tag = "eye" Then
If Screen1 = 10 And Elevator1Open = 1 Then
    If State = 0 Then
        Source.DragIcon = frminv.eye2.DragIcon
    ElseIf State = 1 Then
        Source.DragIcon = frminv.eye1.DragIcon
    ElseIf State = 2 Then
        Source.DragIcon = frminv.eye2.DragIcon
    End If
Else
Source.DragIcon = frminv.eye1.DragIcon
End If
End If
End Sub
Private Sub Imgfront_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'-------Block 1 Foward------------------------------

If Screen1 = 1 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2.jpg")
    Screen1 = 3
ElseIf Screen1 = 2 Then
    If Door1Open = 0 Then
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain1b.jpg")
        Door1Open = 1
        Else
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain1a.jpg")
        Door1Open = 0
    End If

'-------Block 2 Foward------------------------------

ElseIf Screen1 = 3 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain3.jpg")
    Screen1 = 5

'-------Block 2 Backward----------------------------

ElseIf Screen1 = 4 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain1a.jpg")
    Screen1 = 2

'-------Block 3 Foward------------------------------

ElseIf Screen1 = 5 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4.jpg")
    Screen1 = 7

'-------Block 3 Backward----------------------------

ElseIf Screen1 = 6 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2a.jpg")
    Screen1 = 4

'-------Block 4 Foward------------------------------

ElseIf Screen1 = 7 Then
    Elevatorpic
    Screen1 = 10
    Large
    
'-------Block 4 Backward---------------------------
    
ElseIf Screen1 = 9 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain3a.jpg")
    Screen1 = 6

'-------Block 5 Backward---------------------------
ElseIf Screen1 = 12 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4a.jpg")
    Screen1 = 9
    ButtonClicked = False
    Elevator1Open = 0
    
'-------Welcome 1 Foward---------------------------
ElseIf Screen1 = 8 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1b.jpg")
    Screen1 = 13
    Bottomalign
    
ElseIf Screen1 = 13 Then
    Pipescreen
    Pipesize
    Screen1 = 14

'Button Click
ElseIf Screen1 = 11 Then
    Playsoundclick
    If PipeUsed = False Then
        Elevator1Open = 1
    ElseIf PipeUsed = True Then
        Elevator1Open = 2
    End If
ButtonClicked = True

'Elevator Screen
ElseIf Screen1 = 10 Then
    If Elevator1Open = 0 Then
            frmdialogue.Text1.Text = "I should press the button to my left to call the elevator."

    ElseIf Elevator1Open = 1 Then
        frmdialogue.Text1.Text = "The door is jammed!?"
      
    ElseIf Elevator1Open = 2 Then
        Image1.Picture = LoadPicture(App.Path & "\images\elevatorback.jpg")
        Screen1 = 18
        imgfront.Enabled = False
    End If
    
ElseIf Screen1 = 15 Then
        Image1.Picture = LoadPicture(App.Path & "\images\welcome1.jpg")
        Screen1 = 8
        WelcomeDialogue
        
ElseIf Screen1 = 14 Then
        PipeAV = True
        PipeTaken = True
        Pipeinv
        Image1.Picture = LoadPicture(App.Path & "\images\welcome1d.jpg")
        imgfront.Enabled = False
        Playsoundpipetake
        
ElseIf Screen1 = 21 Then
    If ElevatorDoorsOpen = False Then
        If Button2Clicked = False Then
        Image1.Picture = LoadPicture(App.Path & "\images\hallmain5a.jpg")
        Small
        Screen1 = 12
        End If
    Else
        Image1.Picture = LoadPicture(App.Path & "\images\halltop2.jpg")
        Screen1 = 30
        Small
    End If
'Elevator Button Click
ElseIf Screen1 = 20 Then
    If Button2Clicked = False Then
        Button2Clicked = True
        WavFile = "\sound\unjam.wav"
        PlaySound
        ElevatorControl.Interval = 1200
        ElevatorControl.Enabled = True
    End If
    
ElseIf Screen1 = 32 Then
        Image1.Picture = LoadPicture(App.Path & "\images\elevatorback.jpg")
        Screen1 = 18
        imgfront.Enabled = False

ElseIf Screen1 = 30 Then
        frmdialogue.Text1.Text = "This is the end of this demo, Thank you for trying it out!"
End If
'TOGGLE THIS FOR DEBUG MODE
'frmdialogue.Text1.Text = Screen1

End Sub

Private Sub imgfront_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Screen1 = 10 Then
    frminv.statbar1.SimpleText = "Elevator Door"
ElseIf Screen1 = 11 Then
    frminv.statbar1.SimpleText = "Elevator Button"
ElseIf Screen1 = 14 And PipeTaken = False Then
    frminv.statbar1.SimpleText = "A pipe"
ElseIf Screen1 = 20 Then
    frminv.statbar1.SimpleText = "Inside Elevator Button"
Else
    frminv.statbar1.SimpleText = ""
End If

End Sub

Private Sub imgleft_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If Source.Tag = "pipe" Then Source.DragIcon = frminv.imgpipe2.DragIcon
End Sub

Private Sub imgleft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'-------Block 1------------------------------------

If Screen1 = 1 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain1a.jpg")
    Screen1 = 2
    Large

ElseIf Screen1 = 2 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain1.jpg")
    Screen1 = 1
    Door1Open = 0
    Small

'-------Block 2------------------------------------

ElseIf Screen1 = 3 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2a.jpg")
    Screen1 = 4
    
ElseIf Screen1 = 4 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2.jpg")
    Screen1 = 3

'-------Block 3------------------------------------

ElseIf Screen1 = 5 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1.jpg")
    Screen1 = 8
    WelcomeDialogue
    
ElseIf Screen1 = 6 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain3.jpg")
    Screen1 = 5

'-------Block 4------------------------------------

ElseIf Screen1 = 7 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1.jpg")
    Screen1 = 8
    WelcomeDialogue
    
ElseIf Screen1 = 8 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4a.jpg")
    Screen1 = 9
    Small

ElseIf Screen1 = 9 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4.jpg")
    Screen1 = 7

'-------Block 5------------------------------------

ElseIf Screen1 = 10 Then
    Image1.Picture = LoadPicture(App.Path & "\images\button1.jpg")
    Screen1 = 11
    Buttonsize

ElseIf Screen1 = 11 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain5a.jpg")
    Small
    Screen1 = 12

ElseIf Screen1 = 12 Then
    Elevatorpic
    Screen1 = 10
    Large

ElseIf Screen1 = 13 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1a.jpg")
    Screen1 = 15
    imgfront.Enabled = True
    Small

ElseIf Screen1 = 15 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4.jpg")
    Screen1 = 7
    
ElseIf Screen1 = 14 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1a.jpg")
    Screen1 = 15
    imgfront.Enabled = True
    Small

'Elevator Spin

ElseIf Screen1 = 18 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorbutton.jpg")
    Screen1 = 20
    Buttonsize2
    imgfront.Enabled = True
    ElevatorLight.Left = 1890

ElseIf Screen1 = 20 Then
    Elevatorinsidepic
    Screen1 = 21
    Large
    imgfront.Enabled = True
    ElevatorLight.Visible = False
    
ElseIf Screen1 = 21 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorright.jpg")
    Screen1 = 19
    imgfront.Enabled = False
    ElevatorLight.Left = 1910

ElseIf Screen1 = 19 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorback.jpg")
    Screen1 = 18
    imgfront.Enabled = False
    ElevatorLight.Left = 1800

ElseIf Screen1 = 30 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2b.jpg")
    Screen1 = 31
    imgfront.Enabled = False

ElseIf Screen1 = 31 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2a.jpg")
    Screen1 = 32
    imgfront.Enabled = True

ElseIf Screen1 = 32 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2c.jpg")
    Screen1 = 33
    imgfront.Enabled = False

ElseIf Screen1 = 33 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2.jpg")
    Screen1 = 30
    imgfront.Enabled = True

End If

'TOGGLE THIS FOR DEBUG MODE
'frmdialogue.Text1.Text = Screen1
End Sub

Private Sub imgright_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If Source.Tag = "pipe" Then Source.DragIcon = frminv.imgpipe2.DragIcon
End Sub

Private Sub imgright_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'-------Block 1------------------------------------

If Screen1 = 1 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain1a.jpg")
    Screen1 = 2
    Large
ElseIf Screen1 = 2 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain1.jpg")
    Screen1 = 1
    Door1Open = 0
    Small

'-------Block 2------------------------------------

ElseIf Screen1 = 3 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2a.jpg")
    Screen1 = 4

ElseIf Screen1 = 4 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain2.jpg")
    Screen1 = 3

'-------Block 3------------------------------------

ElseIf Screen1 = 5 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain3a.jpg")
    Screen1 = 6

ElseIf Screen1 = 6 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain3.jpg")
    Screen1 = 5

'-------Block 4------------------------------------

ElseIf Screen1 = 7 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4a.jpg")
    Screen1 = 9
    
ElseIf Screen1 = 9 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1.jpg")
    Screen1 = 8
    WelcomeDialogue
    
ElseIf Screen1 = 8 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4.jpg")
    Screen1 = 7
    Small
'-------Block 5------------------------------------

ElseIf Screen1 = 10 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain5a.jpg")
    Screen1 = 12
    Small
    
ElseIf Screen1 = 12 Then
    Image1.Picture = LoadPicture(App.Path & "\images\button1.jpg")
    Buttonsize
    Screen1 = 11

ElseIf Screen1 = 11 Then
    Elevatorpic
    Screen1 = 10
    Large

'-------Welcome 2 spin-----------------------------
ElseIf Screen1 = 13 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1a.jpg")
    Screen1 = 15
    imgfront.Enabled = True
    Small

ElseIf Screen1 = 15 Then
    Image1.Picture = LoadPicture(App.Path & "\images\hallmain4a.jpg")
    Screen1 = 9
    
ElseIf Screen1 = 14 Then
    Image1.Picture = LoadPicture(App.Path & "\images\welcome1a.jpg")
    Screen1 = 15
    imgfront.Enabled = True
    Small

'Elevator spin

ElseIf Screen1 = 18 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorright.jpg")
    Screen1 = 19
    imgfront.Enabled = False
    ElevatorLight.Left = 1910
    
ElseIf Screen1 = 19 Then
    Elevatorinsidepic
    Screen1 = 21
    Large
    imgfront.Enabled = True
    ElevatorLight.Visible = False
    
ElseIf Screen1 = 21 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorbutton.jpg")
    Screen1 = 20
    Buttonsize2
    imgfront.Enabled = True
    ElevatorLight.Left = 1890
    
ElseIf Screen1 = 20 Then
    Image1.Picture = LoadPicture(App.Path & "\images\elevatorback.jpg")
    Screen1 = 18
    imgfront.Enabled = False
    ElevatorLight.Left = 1800
    
ElseIf Screen1 = 30 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2c.jpg")
    Screen1 = 33
    imgfront.Enabled = False
    
ElseIf Screen1 = 33 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2a.jpg")
    Screen1 = 32
    imgfront.Enabled = True
    
ElseIf Screen1 = 32 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2b.jpg")
    Screen1 = 31
    imgfront.Enabled = False

ElseIf Screen1 = 31 Then
    Image1.Picture = LoadPicture(App.Path & "\images\halltop2.jpg")
    Screen1 = 30
    imgfront.Enabled = True
   
End If
'TOGGLE THIS FOR DEBUG MODE
'frmdialogue.Text1.Text = Screen1

End Sub

Private Sub mnucdmusic_Click()
OpenCDPlayer
End Sub

Private Sub mnuexit_Click()
WavFile = ""
PlaySoundLoop
Unload frmimage
frmimage.Visible = False
End Sub

Private Sub mnumusic_Click()
If mnumusic.Checked = False Then
MusicChange
Else
MusicQuit
End If

End Sub

Private Sub mnurestore_Click()
frmimage.Left = 2610
frmimage.Top = 1065
frmdialogue.Left = 1065
frmdialogue.Top = 1065
frminv.Left = 2610
frminv.Top = 5325
End Sub
Private Sub mnurestoresaved_Click()
INIGet
End Sub

Private Sub mnusaved_Click()
Load gamecodes
gamecodes.Show vbModal
End Sub

Private Sub mnusoundon_Click()
    If mnusoundon.Checked = True Then
    SoundQuit
    Else
    SoundChange
End If
End Sub
Private Sub mnustds9_Click()
If mnustds9.Checked = False Then
SongMnuClicked 9
End If
End Sub
Private Sub mnusti_Click()
If mnusti.Checked = False Then
SongMnuClicked 1
End If
End Sub
Private Sub mnustii_Click()
If mnustii.Checked = False Then
SongMnuClicked 2
End If
End Sub
Private Sub mnustiii_Click()
If mnustiii.Checked = False Then
SongMnuClicked 3
End If
End Sub
Private Sub mnustiv_Click()
If mnustiv.Checked = False Then
SongMnuClicked 4
End If
End Sub
Private Sub mnusttng_Click()
If mnusttng.Checked = False Then
SongMnuClicked 5
End If
End Sub
Private Sub mnusttos_Click()
If mnusttos.Checked = False Then
SongMnuClicked 11
End If
End Sub
Private Sub mnustvi_Click()
If mnustvi.Checked = False Then
SongMnuClicked 6
End If
End Sub
Private Sub mnustvii_Click()
If mnustvii.Checked = False Then
SongMnuClicked 7
End If
End Sub
Private Sub mnustviii_Click()
If mnustviii.Checked = False Then
SongMnuClicked 8
End If
End Sub
Private Sub mnustvoy_Click()
If mnustvoy.Checked = False Then
SongMnuClicked 10
End If
End Sub
Private Sub mnuwindowsave_Click()
INIMake
End Sub

Private Sub Timer1_Timer()
MusicVarChange
End Sub
