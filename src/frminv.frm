VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frminv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventory"
   ClientHeight    =   705
   ClientLeft      =   2655
   ClientTop       =   5610
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.StatusBar statbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Status Bar"
      Top             =   405
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4947
            MinWidth        =   4233
            Text            =   "Star Fleet Academy"
            TextSave        =   "Star Fleet Academy"
            Object.ToolTipText     =   "This bar tells you where you are currently located"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3430
            MinWidth        =   2716
            Text            =   "Welcome Center"
            TextSave        =   "Welcome Center"
            Object.ToolTipText     =   "This bar will change depending on what your mouse is over"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3840
      Top             =   0
   End
   Begin VB.Image eye 
      DragIcon        =   "frminv.frx":0000
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   4320
      Picture         =   "frminv.frx":030A
      Tag             =   "eye"
      ToolTipText     =   "Right click on the eye to get help"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image eye1 
      DragIcon        =   "frminv.frx":0614
      Height          =   480
      Left            =   4320
      Picture         =   "frminv.frx":091E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgpipe 
      DragIcon        =   "frminv.frx":0C28
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   0
      OLEDragMode     =   1  'Automatic
      Picture         =   "frminv.frx":0F32
      Tag             =   "pipe"
      ToolTipText     =   "This is a pipe"
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgpipe2 
      DragIcon        =   "frminv.frx":123C
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   0
      OLEDragMode     =   1  'Automatic
      Picture         =   "frminv.frx":1546
      Tag             =   "pipe"
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image eye2 
      DragIcon        =   "frminv.frx":1850
      Height          =   480
      Left            =   4320
      Picture         =   "frminv.frx":1B5A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image eyel 
      DragIcon        =   "frminv.frx":1E64
      Height          =   480
      Left            =   4320
      Picture         =   "frminv.frx":216E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image eyer 
      DragIcon        =   "frminv.frx":2478
      Height          =   480
      Left            =   4320
      Picture         =   "frminv.frx":2782
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frminv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim eyestate As Integer
Dim rannum As Integer

Private Sub eye_Click()
messagebox = MsgBox("This is your eye.  Drag the icons in your inventory over this eye to find out more about them.  You can also drag the eye onto parts of the screen to examine them.", vbInformation, "Your Eye")
End Sub

Private Sub eye_DragDrop(Source As Control, x As Single, y As Single)
If Source.Tag = "pipe" Then
messagebox = MsgBox("This apprears to be a pipe.  I found it in the debris in the Welcome Center of StarFleet Academy.", vbInformation, "A Pipe")
Source.DragIcon = imgpipe2.DragIcon
End If
eye.Picture = eye1.Picture
End Sub
Private Sub eye_DragOver(Source As Control, x As Single, y As Single, State As Integer)

If State = 0 Then
    If Source.Tag = "pipe" Then
        Source.DragIcon = imgpipe.Picture
    End If
    eye.Picture = eye2.Picture
ElseIf State = 1 Then
    If Source.Tag = "pipe" Then
        Source.DragIcon = imgpipe2.DragIcon
    End If
    eye.Picture = eye1.Picture
ElseIf State = 2 Then
    If Source.Tag = "pipe" Then
        Source.DragIcon = imgpipe.Picture
    End If
    eye.Picture = eye2.Picture
End If


End Sub

Private Sub eye_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
statbar1.SimpleText = "Your Eye"
End Sub

Private Sub Form_Load()
'frminv.Width = frmimage.Width
Randomize
eyestate = 1
Timer1.Interval = Int(Rnd * 25000)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
statbar1.SimpleText = "Inventory Window"
End Sub

Private Sub imgpipe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
statbar1.SimpleText = "Pipe"
End Sub

Private Sub statbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
statbar1.SimpleText = "Status Bar"
End Sub
Private Sub Timer1_Timer()
If eyestate = 1 Then
    rannum = Int(Rnd * 10)
        If rannum > 5 Then
            eye.Picture = eye2.Picture
            eyestate = 3
            Timer1.Interval = 600
        Else
            eyestate = 2
            eye.Picture = eye2.Picture
            Timer1.Interval = 800
        End If
        
ElseIf eyestate = 2 Then
    eye.Picture = eye1.Picture
    eyestate = 1
    Timer1.Interval = Int(Rnd * 25000)

ElseIf eyestate = 3 Then
    eye.Picture = eyel.Picture
    eyestate = 4
    Timer1.Interval = 400

ElseIf eyestate = 4 Then
    eye.Picture = eyer.Picture
    eyestate = 1
    Timer1.Interval = 400
    
End If
End Sub
