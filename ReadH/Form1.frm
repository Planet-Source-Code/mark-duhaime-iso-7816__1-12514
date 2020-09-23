VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7875
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7440
      MaskColor       =   &H00808000&
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox sendList 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2985
      Left            =   3920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   3920
   End
   Begin VB.ListBox XPLList 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   3920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open Holes"
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   4680
      TabIndex        =   33
      Top             =   600
      Width           =   2655
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "09 Hole"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   1320
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "E3 Hole"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   840
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "2D Hole"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox USWDec 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox IrdNoDec 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox CardNoDec 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtStat 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800080&
      Height          =   300
      Left            =   2520
      TabIndex        =   25
      Top             =   6020
      Visible         =   0   'False
      Width           =   3865
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   6480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   6000
   End
   Begin VB.TextBox Updates 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3100
      Width           =   1215
   End
   Begin VB.TextBox GuideByte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2660
      Width           =   1215
   End
   Begin VB.TextBox TimeZone 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2200
      Width           =   1215
   End
   Begin VB.TextBox CardStat 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox IRDNo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1280
      Width           =   1215
   End
   Begin VB.TextBox CardNo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   820
      Width           =   1215
   End
   Begin VB.TextBox Messages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   6780
      Width           =   8000
   End
   Begin VB.CommandButton Read 
      Caption         =   "&READ"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox vbCommPort 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6360
      List            =   "Form1.frx":0010
      TabIndex        =   2
      Text            =   "Comm Port"
      Top             =   120
      Width           =   1335
   End
   Begin MSCommLib.MSComm Comm 
      Left            =   120
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   6000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   7900
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00400000&
      X1              =   0
      X2              =   7900
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "TEST HU"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   47
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "FIX 745 MSG."
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   45
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "REPAIR EEPROM"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   44
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "CLOSE OPEN HOLES"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   43
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Receive Data"
      Height          =   195
      Left            =   3940
      TabIndex        =   41
      Top             =   400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "SAVE PACKET LOG"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   39
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Send Data"
      Height          =   200
      Left            =   120
      TabIndex        =   38
      Top             =   400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "SHOW PACKET RESULTS"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   37
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "RUN XPL FILE"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   36
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   7680
      Y1              =   3620
      Y2              =   3620
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "WRITE EEPROM"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   32
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "READ VIA E3"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   29
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "ADVANCED MENU"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   28
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "READ VIA 2D"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   26
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   6045
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSNd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1080
      TabIndex        =   23
      Top             =   6000
      Width           =   525
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "UNMARRY CARD"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "2D HOLE"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "CHECK FOR OPEN"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "EXIT"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   18
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "RESET"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "USW:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   3140
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Guide Byte:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Zone:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Card Status:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "IRD #:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Card #:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label txtATR 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   5500
   End
   Begin VB.Label lblATR 
      Caption         =   "ATR"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   150
      Width           =   495
   End
   Begin VB.Menu mnuBlank1 
      Caption         =   " "
   End
   Begin VB.Menu mnuBlank2 
      Caption         =   " "
   End
   Begin VB.Menu mnublank3 
      Caption         =   "  "
   End
   Begin VB.Menu mnublank4 
      Caption         =   " "
   End
   Begin VB.Menu mnuBlank5 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank6 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank7 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank8 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank9 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank10 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank11 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank12 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank13 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank14 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank15 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank16 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank17 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank18 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank19 
      Caption         =   ""
   End
   Begin VB.Menu mnuBlank20 
      Caption         =   ""
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEEProm 
         Caption         =   "&Important EEProm Locations"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CommPort As Integer
Dim Timeout2 As Integer
Dim Command58 As String
Dim Command2A As String
Dim CardType As String
Dim CardNumber As String
Dim CardStatus As String
Dim IRDNumber As String
Dim TimZone As String
Dim Guide As String
Dim USW As String
Dim StartFlag As Boolean
Dim CardFlag As Boolean
Dim Position As Integer
Dim EEProm As String
Dim ReadAddress As String
Dim ReadOne As Boolean
Dim NumToRead As Integer
Dim EStart As Integer
Dim ESecond As Integer
Dim InTime As Long
Dim DontReceive As Boolean
Dim DontHide As Boolean

Dim Form As Form

Private Sub CardNo_GotFocus()
    If CardFlag Then
        Read.Value = True
        CardFlag = False
    End If
End Sub

Private Sub Comm_OnComm()
    Static Buffer As String
    Dim I As Long
    Dim J As Integer
    Dim Temp As String
    Dim Ev As Integer
    
    Ev = Comm.CommEvent
    'wait event, don't want buffer mixed up
    If Ev = 2 Then
       'For J = 1 To OutByte
            For I = 1 To InTime
    
            Next I
       'Next J
        
        StateChanged = 1
        ' Always buffer incoming data no matter
        ' what generated the OnComm event.
        Buffer = Buffer & Comm.Input
        'Check for P3 first
        If (CardType = "P3") Or (CardType = "") Then
            If Len(Buffer) >= MaxP3Limit Then
                InBuf = Right$(Buffer, MaxP3Limit)
                Buffer = "" ' Right$(Buffer, Len(Buffer) - MaxP3Limit)
            ElseIf Len(Buffer) >= MaxP2Limit Then
                ' Call procedure to Process the received information.
                InBuf = Right$(Buffer, MaxP2Limit)
                Buffer = "" 'Right$(Buffer, Len(Buffer) - MaxP2Limit)
            Else
                InBuf = InBuf + Left$(Buffer, BufLen)
                Buffer = ""
            End If
        ElseIf (CardType = "P2") Or (CardType = "") Then
            ' Check if buffer has grown larger than limit.
            If (MaxP2Limit = 519) And (Len(Buffer) >= 519) Then
                InBuf = Left$(Buffer, MaxP2Limit)
                Buffer = "" 'Right$(Buffer, Len(Buffer) - MaxP2Limit)
            ElseIf Len(Buffer) >= MaxP2Limit Then
                ' Call procedure to Process the received information.
                InBuf = Right$(Buffer, MaxP2Limit)
                Buffer = "" 'Right$(Buffer, Len(Buffer) - MaxP2Limit)
            Else
                InBuf = InBuf + Left$(Buffer, BufLen) '""
                Buffer = ""
            End If
        End If
    ElseIf DontReceive = False Then
        Ev = 0
        StateChanged = 1
    End If
    
End Sub

'***********************
'
'   Exit Routine
'
'***********************
Private Sub Exit_Click()
    If Comm.PortOpen = True Then
        Comm.PortOpen = False
    End If
    Unload Form2
    Unload Me
    Set Form = Nothing
    End
End Sub

Private Sub Command1_Click()
    
    txtStat.Text = "Checking for open holes..."
    ReadAddress = " 8F 20"
    NumToRead = 1
    StartFlag = True
    ReadOne = True
    ReadEEprom
    
    If Right$(InBuf, 3) = "09 " Then
        txtStat.Text = "09 Hole not open."
        Label11.Enabled = False
    Else
        txtStat.Text = "09 Hole is open"
        Label11.Enabled = True
        Label11.Visible = True
    End If
    ResetATR
    InBuf = ""
    ReadAddress = " 85 90"
    ReadEEprom
    If Mid$(InBuf, 4, 2) = "80" Then
        txtStat.Text = "E3 Hole is open."
        Label8.Enabled = True
        Label8.Visible = True
    Else
        txtStat.Text = "E3 Hole not open."
        Label8.Enabled = False
    End If
    InBuf = ""
    EnableControls
    PBar.Value = 0
        
        
End Sub

Private Sub Form_Load()
    Dim Hdl As Integer
    Dim Temp As String
    Dim Msg As String
    
    Dim I As Integer
    Start = 1
    On Error GoTo Handler
    Msg = "Binchk.txt file is missing."
    
    Hdl = FreeFile
    Open "c:\Magic\binchk.txt" For Input As #Hdl
    Msg = "Binchk.txt file is corrupt."
    For I = 1 To 256
        Line Input #Hdl, Temp
        Temp = Left$(Temp, 47)
        Form2.BinList.AddItem Temp
    Next I
    Close #Hdl
    Me.Caption = MagiString + " © 2000 EMIsoft"
    CardFlag = True
    CardType = ""
    StateChanged = 0
    StartFlag = False
    Position = 0
    DontReceive = True
    ByteDelay = ByteMin
    InTime = ByteDelay * 15
    OutByte = 20
    Setupvars
    BufLen = 1
    SendEEprom = False
    ShowPack = False
    Exit Sub
    
Handler:
    MsgBox Msg, vbCritical, "Exiting program...."
    Unload Me
    Set Form = Nothing
    End
End Sub

Private Sub Label_Click(Index As Integer)
    Dim I As Integer
    Dim Temp As String
        
    DisableControls
    sendList.Visible = False
    XPLList.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    
    txtStat.Text = ""
    Select Case Index
        Case 0
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            CardType = ""
            Read.Value = True
        Case 1
            Exit_Click
        Case 2
            StartFlag = True
            Command1.Value = True
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 3
            StartFlag = True
            Hole2D
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 4
            StartFlag = True
            Unmarry
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 5
            StartFlag = True
            ReadOne = False
            EStart = &H80
            ESecond = 0
            ReadAddress = Chr$(32) + Hex(EStart) + Chr$(32) + Hex$(ESecond)
            If ESecond = 0 Then
                ReadAddress = ReadAddress + "0"
            End If
            NumToRead = 16
            ReadEEprom
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 6
            Me.Hide
            Form2.Show vbModal
            Me.Show
        Case 7
            ResetATR
            ReadE3
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 8
            WriteImage
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 9
            SendEEprom = False
            LoadXPL
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Label(10).Enabled = False
            Read.Value = True
            Label(11).Enabled = True
            Label(10).Enabled = True
        Case 10
            If Left$(Label(10).Caption, 4) = "SHOW" Then
                ShowList
            Else
                HideList
            End If
            DontHide = True

        Case 11
            SaveLog
            DontHide = True
        Case 12
            HUTest
        Case 13
            Call CheckOpen
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 14
            Call SafetyCheck
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
        Case 15
            Call Fix745
            If Comm.PortOpen = True Then
                Comm.PortOpen = False
            End If
            Read.Value = True
    End Select
    If Err = 0 Then
        EnableControls
    End If
    
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim I As Integer
    
    Label(Index).ForeColor = &H80

    For I = 0 To 15
        If I <> Index Then
            If Label(I).ForeColor = &H80 Then
                Label(I).ForeColor = &H800000
            End If
        End If
    Next I
End Sub

Private Sub mnuAbout_Click()
    'Form4Start = True
    'Form4.Show 1
    'Unload Form4
    frmCredits.Show 1
    Unload frmCredits
End Sub

Private Sub mnuchkforopen_Click()
    Command1.Value = True
End Sub

Private Sub mnuEEProm_Click()
    Form4Start = True
    frmCredits.Show 1
    Form4Start = False
    Unload frmCredits
End Sub

Private Sub Read_Click()
    Dim Temp As String
    Dim Buf As String
    Dim Count As Integer
    Dim I As Integer
    Static Card As Long
    Dim CardDec As String
    Static TempCard As Long
    Dim CardTemp As Long
    Dim TempIrd As Long
    Dim IRDDec As String
    Dim IRDTemp As Long
    Dim Tempusw As Long
    Dim USWTemp As Long
    Dim UpDec As String
    
    ClearAllControls
    On Error GoTo SendEnd
    PBar.Max = 100
    txtATR.Caption = ""
    Temp = vbCommPort.Text
    If Temp = "Comm Port" Then
        Temp = "Comm 1"
    End If
    CommPort = Val(Right$(Temp, 1))
    Comm.CommPort = CommPort
    If Comm.PortOpen = False Then
        Comm.Settings = "9600,O,8,2"
        Comm.PortOpen = True
    End If
    On Error Resume Next
    
    DisplayMessage "Attempting Answer to Reset."
    vbCommPort.Visible = False
    lblATR.Visible = True
    txtATR.Visible = True
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 200
    Comm.RTSEnable = True
    Comm.DTREnable = False
    Count = 0
Top:
    InBuf = ""
    Comm.RThreshold = 1
    Comm.InputLen = 0
    MaxP2Limit = 13
    MaxP3Limit = 20
    Timer1.Enabled = True
    Timer1.Interval = 200
    Do While TimeOut = 0
        DoEvents
    Loop
    StateChanged = 0
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Enabled = True
    Comm.RTSEnable = False

    'rem P3 card
    Do While Len(InBuf) < 16
        If TimeOut = 1 Then
            'rem P2 card
            If Len(InBuf) > 11 Then
                'rem add here for timing interval checks
                Exit Do
            Else
                If ByteDelay >= ByteMax Then
                    DisplayMessage "Timeout on Reset. Check cables, connection, comm port."
                    vbCommPort.Visible = True
                    lblATR.Visible = False
                    txtATR.Visible = False
                    Beep
                    ByteDelay = ByteMin
                    OutByte = 20
                    Err = 1
                    GoTo SendEnd
                Else
                    ByteDelay = ByteDelay + ByteMin
                    InTime = ByteDelay * 15
                    OutByte = OutByte + 1
                End If
            End If
        End If
        DoEvents
    Loop
    StateChanged = 0
    Temp = ""
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
        If I = 3 Then
            Mask = HoldAtr And &HF
        End If
    Next I
    If Len(Temp) >= 60 Then
        Temp = Left$(Temp, 60)
        If Temp <> "3F 7F 13 25 03 38 B0 04 FF FF 4A 50 00 00 29 48 55 55 00 00 " Then
            DisplayMessage "Invalid ATR for P3..."
            Count = Count + 1
            If Count < 5 Then
                GoTo Top
            Else
                Beep
                GoTo SendEnd
            End If
        End If
        DisplayMessage "Valid P3 Access Card inserted."
        txtATR.Caption = Temp
        CardType = "P3"
    ElseIf Len(Temp) = 39 Then
        If Temp <> "3F 78 12 25 01 40 B0 03 4A 50 20 48 55 " Then
            DisplayMessage "Invalid ATR for P2..."
            Count = Count + 1
            If Count < 5 Then
                GoTo Top
            Else
                Beep
                GoTo SendEnd
            End If
        End If
        DisplayMessage "Valid P2 Access Card inserted."
        CardType = "P2"
        txtATR.Caption = Temp
    ElseIf InStr(1, Temp, "99 99 99 99 99 99") Then
        'rem 99 check etc.
        DisplayMessage "Card is 99'ed and requires unloop."
        txtATR.Caption = Temp
        GoTo SendEnd
    Else
        DisplayMessage "Please insert a valid access card."
        Count = Count + 1
        If Count < 5 Then
            GoTo Top
        Else
            Beep
            GoTo SendEnd
        End If
    End If
    
    InTime = ByteDelay * 17
    
    Position = 100 / 13
    PBar.Value = Position
    
    InBuf = ""
    Select Case Mask
        Case 1
            Comm.Settings = "9600,O,8,2"
        Case 2
            Comm.Settings = "19200,O,8,2"
        Case 3
            Comm.Settings = "38400,O,8,2"
        Case 5
            Comm.Settings = "115200,O,8,2"
    End Select
    
    Select Case CardType
        Case "P2", "P3"
        
            TwoA
            Position = 65
            PBar.Value = Position
            Comm.InputLen = 0
            Temp = Comm.Input
                                
            InvertData
            Command2A = ReturnData
            
            If Left$(Command2A, 5) <> "80 2A" Then
                DisplayMessage "Card did not respond to 2A command."
                GoTo SendEnd
                InBuf = ""
            End If
            
            TempAtr = Asc(Mid$(InBuf, 23, 1))
            ConvertAtr
            TempCard = CLng(HoldAtr) * 256 * 256 * 256
            Card = TempCard
            TempCard = 0
            TempAtr = Asc(Mid$(InBuf, 24, 1))
            ConvertAtr
            TempCard = CLng(HoldAtr) * 256 * 256
            Card = Card + TempCard
            TempCard = 0
            TempAtr = Asc(Mid$(InBuf, 25, 1))
            ConvertAtr
            TempCard = CLng(HoldAtr) * 256
            Card = Card + TempCard
            TempCard = 0
            TempAtr = Asc(Mid$(InBuf, 26, 1))
            ConvertAtr
            Card = Card + CLng(HoldAtr)
            CardDec = Str(Card) + "_"
            TempAtr = Asc(Mid$(InBuf, 27, 1))
            ConvertAtr
            TempIrd = CLng(HoldAtr) * 256 * 256 * 256
            IRDTemp = TempIrd
            TempIrd = 0
            TempAtr = Asc(Mid$(InBuf, 28, 1))
            ConvertAtr
            TempIrd = CLng(HoldAtr) * 256 * 256
            IRDTemp = IRDTemp + TempIrd
            TempIrd = 0
            TempAtr = Asc(Mid$(InBuf, 29, 1))
            ConvertAtr
            TempIrd = CLng(HoldAtr) * 256
            IRDTemp = IRDTemp + TempIrd
            TempAtr = Asc(Mid$(InBuf, 30, 1))
            ConvertAtr
            IRDTemp = IRDTemp + CLng(HoldAtr)
            IRDDec = Str$(IRDTemp) + "_"
            TempAtr = Asc(Mid$(InBuf, 31, 1))
            ConvertAtr
            Tempusw = CLng(HoldAtr) * 256
            TempAtr = Asc(Mid$(InBuf, 32, 1))
            ConvertAtr
            Tempusw = Tempusw + CLng(HoldAtr)
            UpDec = Str$(Tempusw)
            
            InBuf = ""
            FiftyEight
            Position = 100
            PBar.Value = Position
            InvertData
            Command58 = ReturnData
            
            Comm.InputLen = 0
            Temp = Comm.Input
            If Left$(Command58, 5) <> "17 58" Then
                DisplayMessage "Card did not respond to 58 command."
                GoTo SendEnd
            End If
            
    End Select
    
    Select Case CardType

    
        Case "P2", "P3"
            
            CardNumber = Mid$(Command2A, 67, 2) + Mid$(Command2A, 70, 2) + Mid$(Command2A, 73, 2) + Mid$(Command2A, 76, 2)
            CardNo.Text = CardNumber
            CardNoDec.Text = CardDec
            
            IRDNumber = Mid$(Command2A, 79, 2) + Mid$(Command2A, 82, 2) + Mid$(Command2A, 85, 2) + Mid$(Command2A, 88, 2)
            IRDNo.Text = IRDNumber
            IrdNoDec.Text = IRDDec
            USW = Mid$(Command2A, 91, 2) + Mid$(Command2A, 94, 2)
            Updates.Text = USW
            USWDec.Text = UpDec
            CardStatus = Mid$(Command58, 7, 2)
            TimZone = Mid$(Command58, 37, 2)
            Guide = CardStatus + " " + Mid$(Command58, 43, 2)
            
            Select Case CardStatus
                Case "00"
                    If CardNumber = IRDNumber Then
                        CardStat = "Off - Not Married"
                        Label(4).Enabled = False
                    Else
                        CardStat = "Off - Married"
                        Label(4).Enabled = True
                    End If
                Case "20"
                    If CardNumber = IRDNumber Then
                        CardStat = "Off - Not Married"
                        Label(4).Enabled = False
                    Else
                        CardStat = "Off - Married"
                        Label(4).Enabled = True
                    End If
                Case "05"
                    If CardNumber = IRDNumber Then
                        CardStat = "On - Not Married"
                        Label(4).Enabled = False
                    Else
                        CardStat = "On - Married"
                        Label(4).Enabled = True
                    End If
                Case "25"
                    If CardNumber = IRDNumber Then
                        CardStat = "On - Not Married"
                        Label(4).Enabled = False
                    Else
                        CardStat.Text = "On - Married"
                        Label(4).Enabled = True
                    End If
            End Select
            Select Case TimZone
                Case "A0"
                    TimeZone.Text = "Pacific"
                Case "A2"
                    TimeZone.Text = "Mountain"
                Case "A4"
                    TimeZone.Text = "Central"
                Case "A6"
                    TimeZone.Text = "Eastern"
                Case Else
                    TimeZone.Text = TimZone
            End Select
            GuideByte.Text = Guide
    End Select
    
SendEnd:
    If Err <> 0 Then
        Beep
        DisplayMessage "Unable to open comm port."
        ByteDelay = ByteMin
        OutByte = 20
        
    End If
    
    
    Position = 0
    PBar.Value = Position
    If Err = 0 Then
        Label(3).Enabled = True
        Label(2).Enabled = True
        Label(5).Enabled = True
        Label(6).Enabled = True
        Label(7).Enabled = True
        Label(8).Enabled = True
        Label(9).Enabled = True
        Label(13).Enabled = True
        Label(14).Enabled = True
        Label(15).Enabled = True
    Else
        Label(3).Enabled = False
        Label(2).Enabled = False
        Label(5).Enabled = False
        Label(6).Enabled = False
        Label(7).Enabled = False
        Label(8).Enabled = False
        Label(9).Enabled = False
        Label(13).Enabled = False
        Label(14).Enabled = False
        Label(15).Enabled = False
    End If
    
    Label(0).Enabled = True
    Label(1).Enabled = True
    mnuHelp.Enabled = True
    If StartFlag Then
        Exit Sub
    End If
    'Comm.RTSEnable = False
    'Comm.DTREnable = True
    'Comm.RTSEnable = False
    If Comm.PortOpen = True Then
        Temp = Comm.Input
    End If
    Comm.InputLen = 0

    Timer1.Enabled = False
    'Comm.PortOpen = False
    If Start = 0 Then
        Start = 1
        Call SafetyCheck
    End If
    If StartFlag = False Then
        StartFlag = True
        'Command1.Value = 1
    End If
    
End Sub

Private Sub DisplayMessage(s As String)
    Messages.Text = s
End Sub

Private Sub Timer1_Timer()
    TimeOut = 1
    Timer1.Enabled = False
End Sub

Sub TwoA()
    Dim Temp As String

    DontReceive = True
    MaxP2Limit = 132
    MaxP3Limit = 132
    
    Temp = Comm.Input
    InBuf = ""
    Temp = ""
    TempAtr = &H48
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp
    
    TempAtr = &H2A
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp
    
    TempAtr = 0
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp

    SendChar Temp
    
    If CardType = "P2" Then
        TempAtr = &H80
    ElseIf CardType = "P3" Then
        TempAtr = &H80
    End If
    ConvertAtr
    DontReceive = True
    
    Temp = Chr(HoldAtr)
    SendChar Temp
End Sub
Sub FiftyEight()
    Dim Temp As String
    
    DontReceive = True
    InBuf = ""
    Temp = ""
    TempAtr = &H48
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp
    
    TempAtr = &H58
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp
    
    TempAtr = 0
    ConvertAtr
    Temp = Chr(HoldAtr)
    SendChar Temp
    SendChar Temp
    
    DontReceive = True
    TempAtr = &H17
    ConvertAtr
    Temp = Chr(HoldAtr)
    MaxP2Limit = 27
    MaxP3Limit = 27
    SendChar Temp
End Sub


Sub ClearAllControls()
    
    CardNo.Text = ""
    CardNoDec.Text = ""
    IRDNo.Text = ""
    IrdNoDec.Text = ""
    CardStat.Text = ""
    TimeZone.Text = ""
    GuideByte.Text = ""
    Updates.Text = ""
    USWDec.Text = ""
    
    lblStat.Visible = True
    txtStat.Visible = True
End Sub

Sub InvertData()
    Dim Temp As String
    Dim I As Integer
    
    Temp = ""
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    ReturnData = Temp
End Sub

Private Sub Timer2_Timer()
    Timeout2 = 1
    Timer2.Enabled = False
End Sub

Private Sub Remove29()
    Dim I As Integer
    Dim Temp As String
    Dim pos As Integer
    Dim pos2 As Integer
    
    XPLList.Clear
    pos = 1
    For I = 1 To 16
        pos2 = InStr(pos, Un29, Chr$(13) + Chr$(10))
        If I = 16 Then
            Temp = Mid$(Un29, pos, 3)
        Else
            Temp = Mid$(Un29, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    RunXPL
    If ReturnData = "90 80 " Then
        txtStat.Text = "Remove 29 successful"
    Else
        txtStat.Text = "Remove 29 unsuccessful"
    End If
    
End Sub

Sub Unmarry()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    Dim Count As Integer
    
    txtStat.Text = ""
    Comm.RTSEnable = False
    DontReceive = True
    InBuf = ""
    Temp = ""
    Temp1 = ""
    InBuf = ""
    Temp = ""
    TempAtr = &H48
    ConvertAtr
    Temp = Chr(HoldAtr)
    
    TempAtr = &H40
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Temp + Chr(HoldAtr)
    
    TempAtr = 0
    ConvertAtr
    Temp = Temp + Chr(HoldAtr)
    
    Temp = Temp + Chr(HoldAtr)
    
    'Temp = Temp + Chr(HoldAtr)
    'Temp = Temp + Chr(HoldAtr)
    TempAtr = &H11
    ConvertAtr
    Temp = Temp + Chr(HoldAtr)
    Position = 15
    PBar.Value = Position
    MaxP2Limit = 6
    DontReceive = True
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    InBuf = ""
    
    DontReceive = True
    TempAtr = &H60
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 200
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &HD5
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    StateChanged = 0
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 200
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H2
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H85
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    StateChanged = 0
    Position = Position + 5
    PBar.Value = Position
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H8E
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    StateChanged = 0
    Position = Position + 5
    PBar.Value = Position
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &HE3
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    StateChanged = 0
    Position = Position + 5
    PBar.Value = Position
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H8
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H4
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H83
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &HD0
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &H0
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    'Temp = Temp + Chr(HoldAtr)
    lblSNd.Caption = Hex$(TempAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    'Temp = Temp + Chr(HoldAtr)
    lblSNd.Caption = Hex$(TempAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    'Temp = Temp + Chr(HoldAtr)
    lblSNd.Caption = Hex$(TempAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    'Temp = Temp + Chr(HoldAtr)
    lblSNd.Caption = Hex$(TempAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    TempAtr = &HBB
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    Position = Position + 5
    PBar.Value = Position
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    
    MaxP2Limit = 19
    DontReceive = True
    TempAtr = &H0
    lblSNd.Caption = Hex$(TempAtr)
    ConvertAtr
    'Temp = Temp + Chr(HoldAtr)
    Temp = Chr(HoldAtr)
    StateChanged = 0
    Comm.Output = Temp
    TimeOut = 0
    Timer1.Enabled = False
    Timer1.Interval = 75
    Timer1.Enabled = True
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    Position = Position + 5
    PBar.Value = Position
    lblSNd.Caption = ""
    Temp = ""
    'convert read
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    
    InBuf = ""
    If InStr(1, Temp, "60 D5 02 85 8E E3 08 04 83 D0 00 00 00 00 00 BB 00 90 00") = False Then
        DisplayMessage "Unmarry card was unsuccessful."
        StartFlag = False
        Temp = Comm.Input
        Exit Sub
    End If
End Sub

Sub Update()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    InBuf = ""
    txtStat.Text = ""
    Temp1 = "48 40 00 00 0D"
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Temp2 + Chr(HoldAtr)
    Next I
    MaxP2Limit = 6
    Position = 12
    PBar.Value = Position
    SendChar Temp2
    Temp = ""
    'convert read
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    InBuf = ""
    If InStr(1, Temp, "48 40 00 00 0D 40 ") = False Then
        txtStat.Text = "Failed to change Updates."
        StartFlag = False
        Temp = Comm.Input
        Exit Sub
    End If
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    DontReceive = True
    
    Temp1 = "60 08 84 00 00 00 00 00 00"
    ' rem send param for usw
    Temp1 = Temp1 + " 00 1A B4 00"
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        lblSNd.Caption = Right$(Temp, 2)
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Chr(HoldAtr)
        Position = Position + 3
        PBar.Value = Position
        StateChanged = 0
        'rem updates position
        If I = 39 Then
            DontReceive = True
        End If
        MaxP2Limit = 89
        Comm.Output = Temp2
        TimeOut = 0
        Timer1.Enabled = False
        Timer1.Interval = 75
        Timer1.Enabled = True
        Do While (TimeOut = 0) Or (StateChanged = 0)
            DoEvents
        Loop
    Next I

    Temp = ""
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    
    InBuf = ""
    
    'Comm.Output = ""
    lblSNd.Caption = ""
    
' well remember to add read code
'to verify write operation
    
    If InStr(1, Temp, "90 80") Then
        txtStat.Text = "Successful"
    Else
            txtStat.Text = "2D Hole was unsuccessful."
            StartFlag = False
            Temp = Comm.Input
            Exit Sub
    End If
    Temp = Comm.Input
    
End Sub

Sub ReadEEprom()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    Dim Count As Integer
    Dim cnt As Integer
    
    DisableControls
    If StartFlag = False Then
        Form2.EEPromList.Clear
        Form2.UpdateList.Clear
    End If
    DontReceive = True
    PBar.Value = 0
    
    txtStat.Text = ""
    Counter = 0
    Init = True
    cnt = 0
    Position = 0
    If NumToRead > 1 Then
        PBar.Max = 84 * NumToRead
    Else
        PBar.Max = 84
    End If
            
    For J = 1 To NumToRead
        Temp = ""
        Temp1 = ""
        Temp2 = ""
        InBuf = ""
        EEProm = ""

        Temp1 = "48 40 00 00 54"
        For I = 1 To Len(Temp1)
            Temp = "&H" + (Mid$(Temp1, I, 2))
            If Left$(Temp, 3) = "&HR" Then
                Exit For
            End If
            I = I + 2
            If I > 12 Then
                DontReceive = True
                BufLen = 2
            Else
                BufLen = 1
            End If
            
            TempAtr = CInt(Temp)
            ConvertAtr
            Temp2 = Chr(HoldAtr)
            MaxP2Limit = 2
            MaxP3Limit = 2
            SendChar Temp2
            
        Next I
        
        PBar.Value = Position
        Temp = ""
        'convert read
        For I = 1 To Len(InBuf)
            TempAtr = Asc(Mid$(InBuf, I, 1))
            ConvertAtr
            If HoldAtr < &H10 Then
                Temp = Temp + "0" + Hex$(HoldAtr) + " "
            Else
                Temp = Temp + Hex$(HoldAtr) + " "
            End If
        Next I
        InBuf = ""
        If InStr(1, Temp, "54 40 ") = False Then
            txtStat.Text = "Failed to read EEProm."
            StartFlag = False
            Temp = Comm.Input
            'Form2.Show vbModal
            Exit Sub
        End If
        Temp = Comm.Input
        Comm.InputLen = 0
        Temp = ""
        Temp1 = ""
        Temp2 = ""
        DontReceive = True
        cnt = 0
        Temp1 = "09 11 00 00 30 60 00 06 39 00 04 F4 22 33 CF 03"
        Temp1 = Temp1 + " 0E 1B 00 CF 03 0E 1B 00 CF 03 0E 1B 00 CF 03 0E"
        Temp1 = Temp1 + " 1B 00 CF 03 0E 1B 00 BB 00 12 00 00 00 00 00 00"
        Temp1 = Temp1 + " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
        Temp1 = Temp1 + " 00 00 84 08 D5 85 46 13 8A 1C 00 00 00 00 00 60 BB 02"

        'rem address here
        Temp1 = Temp1 + ReadAddress
        'StartAddress = ByteOffset + 0x8000;
        For I = 1 To Len(Temp1)
            Temp = "&H" + (Mid$(Temp1, I, 2))
            If Left$(Temp, 3) = "&HR" Then
                Exit For
            End If
            I = I + 2
            cnt = cnt + 1
            lblSNd.Caption = Right$(Temp, 2)
            TempAtr = CInt(Temp)
            ConvertAtr
            Temp2 = Chr(HoldAtr)
            Position = Position + 1
            PBar.Value = Position
            StateChanged = 0
            If I = 252 Then
                DontReceive = True
                BufLen = 3
            Else
                BufLen = 1
            End If
            MaxP2Limit = 259
            MaxP3Limit = 259
            InTime = ByteDelay * 25
            Comm.Output = Temp2
            TimeOut = 0
            Timer1.Enabled = False
            Timer1.Interval = 200
            Timer1.Enabled = True
            Do While (TimeOut = 0) Or (StateChanged = 0)
                DoEvents
            Loop
        Next I
        cnt = 0
        If InBuf <> "" Then
            Label7.Enabled = True
            Label7.Visible = True
            Temp = Comm.Input
            Comm.InputLen = 0
        Else
            txtStat.Text = "Unable to read EEProm."
            StartFlag = False
            Temp = Comm.Input
            InTime = ByteDelay * 17
            Comm.InputLen = 0
            Exit Sub
        End If
        
        Count = 0
        Temp = ""
        
        InBuf = Mid$(InBuf, 2, 256)
        
        For I = 1 To Len(InBuf)
            Count = Count + 1
            TempAtr = Asc(Mid$(InBuf, I, 1))
            ConvertAtr
            If HoldAtr < &H10 Then
                Temp = Temp + "0" + Hex$(HoldAtr) + " "
            Else
                Temp = Temp + Hex$(HoldAtr) + " "
            End If
            If Count = 16 Then
                If StartFlag = False Then
                    Count = 0
                    cnt = cnt + 1
                    Form2.EEPromList.AddItem Temp
                    Form2.UpdateList.AddItem Temp
                    EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
                    Temp = ""
                Else
                    InBuf = Temp
                    Temp = ""
                    Exit Sub
                End If
            End If
        Next I
    
        InBuf = ""
        'Comm.Output = ""
        lblSNd.Caption = ""
        lblStat.Visible = True
        txtStat.Visible = True
    
    Temp = Comm.Input
    'rem set readaddress and status values
    EStart = EStart + 1
    ReadAddress = Chr$(32) + Hex$(EStart) + " 00"
    ResetATR
    Next J
    
    PBar.Value = 0
    Me.Hide
    Form2.Show vbModal
    Me.Show
    txtStat.Text = "Successful"
    Temp = Comm.Input
    E3Flag = False
    Label(7).Visible = True
    'EnableControls
    InTime = ByteDelay * 17
    
End Sub


'rem this gets atr if e3 is open
Sub ReadRom()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    InBuf = ""
    EEProm = ""
    'For Start = 1 To 16
    
    Temp1 = "48 40 00 00 09"
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Temp2 + Chr(HoldAtr)
    Next I
    MaxP2Limit = 6
    Position = 12
    PBar.Value = Position
    SendChar Temp2
    Temp = ""
    'convert read
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    InBuf = ""
    If InStr(1, Temp, "48 40 00 00 09 40 ") = False Then
        txtStat.Text = "Failed to change Updates."
        StartFlag = False
        Temp = Comm.Input
        Exit Sub
    End If
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    DontReceive = True

    Temp1 = "60 D5 02 7F A3 E3 00 "
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        lblSNd.Caption = Right$(Temp, 2)
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Chr(HoldAtr)
        Position = Position + 1
        PBar.Value = Position
        StateChanged = 0
        If I > 18 Then
            DontReceive = True
        End If
        MaxP2Limit = 20
        Comm.Output = Temp2
        TimeOut = 0
        Timer1.Enabled = False
        Timer1.Interval = 75
        Timer1.Enabled = True
        Do While (TimeOut = 0) Or (StateChanged = 0)
            DoEvents
        Loop
        SendChar Temp2
    Next I

    Temp = ""
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    
    InBuf = ""
    EEProm = EEProm + Mid$(Temp, 21, Len(Temp) - 21)
    
    'Comm.Output = ""
    lblSNd.Caption = ""
    lblStat.Visible = True
    txtStat.Visible = True

' well remember to add read code
'to verify write operation
    
    If InStr(1, Temp, "90 80") Then
        txtStat.Text = "Successful"
    Else
            txtStat.Text = "2D Hole was unsuccessful."
            StartFlag = False
            Temp = Comm.Input
            Exit Sub
    End If
    Temp = Comm.Input

End Sub

Sub ReadE3()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    Dim Count As Integer
    Dim cnt As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    
    SendEEprom = True
    DisableControls
    Form2.EEPromList.Clear
    Form2.UpdateList.Clear
    Init = True
    Counter = 0
    
    InTime = ByteDelay * 25
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    If Label8.Visible = True Then
        GoTo E3isOpen
    End If
    
    'check for open E3 hole first
    For I = 1 To 5
        pos2 = InStr(pos, TestE3, Chr$(13) + Chr$(10))
        If I = 5 Then
            Temp = Mid$(TestE3, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(TestE3, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    RunXPL
    If Right$(InBuf, 6) <> "90 80 " Then
    'ok E3 wasn't open...lets open it.
        pos = 1
        For I = 1 To 9
            pos2 = InStr(pos, RdE3, Chr$(13) + Chr$(10))
            If I = 9 Then
                Temp = Mid$(RdE3, pos, 3)
            ElseIf pos2 = 0 Then
                Exit For
            Else
                Temp = Mid$(RdE3, pos, pos2 - pos)
            End If
            XPLList.AddItem Temp
            pos = pos2 + 2
        Next I
        RunXPL
        If Right$(InBuf, 6) <> "90 80 " Then
            E3Flag = False
            GoTo SendEnd
        End If
        E3Flag = True
    Else
        E3Flag = True

    End If

E3isOpen:

    'ok E3 now open
    Label8.Enabled = True
    Label8.Visible = True
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE31, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE31, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE31, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '1st block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    'lets make sure its valid
    If Left$(InBuf, 6) <> "33 14 " Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        If I = 721 Then
            Temp = "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
        End If
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
        
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE32, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE32, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE32, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '2nd block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "02 07 23") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE33, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE33, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE33, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '3nd block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "22 22 22") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE34, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE34, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE34, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '4th block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "15 81 E4") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If

    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE35, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE35, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE35, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '5th block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "11 49 74") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE36, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE36, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE36, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '6th block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "12 05 7C") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE37, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE37, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE37, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '7th block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "20 47 03") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, RdE38, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE38, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE38, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    '8th block 512 bytes
    RunXPL
    InBuf = Right$(InBuf, (512 * 3))
    If InStr(1, InBuf, "E5 2D 24") = False Then
        GoTo SendEnd
    End If
    
    For I = 1 To Len(InBuf) Step 48
        Temp = Mid$(InBuf, I, 48)
        cnt = cnt + 1
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
        EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
        Temp = ""
    Next I
    If Len(InBuf) < 1526 Then
        GoTo SendEnd
    End If
    
    ResetATR
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    pos = 1
    For I = 1 To 4
        pos2 = InStr(pos, RdE39, Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(RdE39, pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(RdE39, pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    
    'clean up 80F0 location used for E3 read
    RunXPL
    InBuf = ""
    SendEEprom = False
    'XPLList.Clear
    'sendList.Clear
    txtStat.Text = "Read eeprom successful."
    InTime = ByteDelay * 17
    SendEEprom = False
    If Form2.EEPromList.ListCount < 255 Then
        GoTo SendEnd
    End If
    Me.Hide
    Form2.Show vbModal
    Me.Show
    Exit Sub
    
SendEnd:
    'only read a portion or none
    'fill remainder with 00's
    Temp = "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    I = Form2.EEPromList.ListCount
    If I = -1 Then
        I = 0
    End If
    For J = I To 255
        Form2.EEPromList.AddItem Temp
        Form2.UpdateList.AddItem Temp
    Next J
    Temp = ""
    'unsuccessful
    InBuf = ""
    SendEEprom = False
    'XPLList.Clear
    'sendList.Clear
    txtStat.Text = "Unable to read eeprom."
    InTime = ByteDelay * 17
    SendEEprom = False
    Me.Hide
    Form2.Show vbModal
    Me.Show
End Sub

Sub DisableControls()

    Label(0).Enabled = False
    Label(1).Enabled = False
    Label(2).Enabled = False
    Label(3).Enabled = False
    Label(4).Enabled = False
    Label(5).Enabled = False
    Label(6).Enabled = False
    Label(7).Enabled = False
    Label(8).Enabled = False
    Label(9).Enabled = False
    Label(10).Enabled = False
    Label(11).Enabled = False
    Label(13).Enabled = False
    Label(14).Enabled = False
    Label(15).Enabled = False
    mnuHelp.Enabled = False
    lblStat.Visible = True
    txtStat.Visible = True
    
End Sub

Sub EnableControls()
    
    Label(0).Enabled = True
    Label(1).Enabled = True
    Label(2).Enabled = True
    Label(3).Enabled = True
    'Label(4).Enabled = True
    Label(5).Enabled = True
    Label(6).Enabled = True
    Label(7).Enabled = True
    Label(8).Enabled = True
    Label(9).Enabled = True
    mnuHelp.Enabled = True
    If DontHide = True Then
        Label(10).Enabled = True
        Label(11).Enabled = True
        DontHide = False
    End If
    Label(13).Enabled = True
    Label(14).Enabled = True
    Label(15).Enabled = True
    
End Sub

Sub LoadXPL()
    Dim Hdl As Integer
    Dim I As Integer
    Dim Temp As String
    Dim J As Integer
    Dim Msg As String
    Dim Hold As Integer
    
    Hdl = FreeFile
    BinFile = ""
    FilName = "C:\Magic\XPLs\"
    Form3.FileList.Pattern = "*.XPL"
    Form3.Show vbModal
    If BinFile = "" Then
        Exit Sub
    End If
        
    XPLList.Clear
    'On Error Resume Next
    
    Open "C:\Magic\XPLs\" + BinFile For Input As #Hdl
    If Err > 0 Then
        Msg = "Error loading XPL file."
        GoTo Handler
    End If
    
    Do
        For I = 1 To LOF(Hdl)
            If EOF(Hdl) Then
                Exit Do
            End If
            Line Input #Hdl, Temp
            XPLList.AddItem Temp
        Next I
    Loop

    Close #Hdl
    RunXPL
    
    Exit Sub
    
Handler:
    
    txtStat.Text = Msg
    Close #Hdl
    
End Sub

Sub SaveLog()
    Dim Hdl As Integer
    Dim I As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim J As Integer
    Dim Msg As String
    Dim Hold As Integer
    
    Hdl = FreeFile
    FilName = "C:\Magic\Logs\"
    Temp = InputBox("Save Present Packet Log as FileName", MagiString + " © 2000 -- Save Packet Log")
    I = InStr(1, Temp, ".")
    If I = 0 Then
        Temp = Temp + ".Log"
    Else
        Temp = Left$(Temp, I - 1) + ".Log"
    End If
    Temp1 = FilName + Temp
    
    Open Temp1 For Output Access Write As #Hdl
    'On Error Resume Next
    sendList.Visible = True
    XPLList.Visible = True
   ' I = sendList.ListCount
   ' For J = 0 To I - 1
   '     sendList.ListIndex = J
   '     Temp = sendList.Text
   '     Write #1, Temp
   ' Next
    Temp = sendList.Text
    Write #1, Temp
    Close #Hdl
    
End Sub

Sub HUTest()
    Dim I As Long
    Dim cnt As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim one As String
    Dim two As String
    Dim three As String
    Dim four As String
    
    SendEEprom = True
    XPLList.Visible = True
    sendList.Visible = True
    
    For I = 33620225 To 2000000000
        If I <> &H3E Then
            Temp1 = Hex(I)
            Temp = "48 42 00 00 04"
            XPLList.Clear
            sendList.Clear
            XPLList.AddItem Temp
            Temp = "R01"
            XPLList.AddItem Temp
            If I < 4096 Then
                one = "00"
                two = "00"
                three = "0" + Left$(Temp1, 1)
                four = Right$(Temp1, 2)
            ElseIf I < 65536 Then
                one = "00"
                two = "0" + Left$(Temp1, 1)
                three = Mid$(Temp1, 2, 2)
                four = Right$(Temp1, 2)
            ElseIf I < 1048576 Then
                one = "00"
                two = Left$(Temp1, 2)
                three = Mid$(Temp1, 3, 2)
                four = Right$(Temp1, 2)
            ElseIf I < 268435456 Then
                one = "0" + Left$(Temp1, 1)
                two = Mid$(Temp1, 2, 2)
                three = Mid$(Temp1, 4, 2)
                four = Right$(Temp1, 2)
            Else
                one = Left$(Temp1, 2)
                two = Mid$(Temp1, 3, 2)
                three = Mid$(Temp1, 5, 2)
                four = Right$(Temp1, 2)
            End If
            
            'two = Temp1
            'Temp = "90 13 40 03 "
            Temp = one + " " + two + " " + three + " " + four
            'Temp = Temp + " 03 85 91 80 08 6b 5c 18"
            XPLList.AddItem Temp
            'Temp = "75 c7 cd 5a e3"
            'XPLList.AddItem Temp
            Temp = "R02"
            XPLList.AddItem Temp
            
            RunXPL
            If (Len(InBuf) > 69) Or (Right$(InBuf, 6) <> "90 00 ") Then
                
            End If
            InBuf = ""
            ResetATR
        End If
    Next I
    SendEEprom = False
    
End Sub

Sub Hole2D()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    Dim Count As Integer
    
    DisableControls
    DontReceive = True
    Count = 0
    InBuf = ""
Top:
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    BufLen = 1
    Temp1 = "48 40 00 00 58"
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        If I = 15 Then
            DontReceive = True
            BufLen = 2
        Else
            BufLen = 1
        End If
        
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Chr(HoldAtr)
        MaxP3Limit = 6
        MaxP2Limit = 6
        SendChar Temp2
    Next I

    Position = 12
    PBar.Value = Position
            
    Temp = ""
    'convert read
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    InBuf = ""
     If InStr(1, Temp, "48 40 00 00 58 40") = False Then
        Count = Count + 1
        If Count > 4 Then
            DisplayMessage "Remove 29 was unsuccessful."
            StartFlag = False
            Temp = Comm.Input
            EnableControls
            Exit Sub
        Else
            Temp = Comm.Input
            GoTo Top
        End If
    End If
    Temp = Comm.Input
    
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    DontReceive = True
    
    Temp1 = Temp1 + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03"
    Temp1 = Temp1 + " 0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e"
    Temp1 = Temp1 + " 1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00"
    Temp1 = Temp1 + " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
    Temp1 = Temp1 + " 00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 "
    '                         85 91 80
    Temp1 = Temp1 + "bb 06 01 85 91 80 00 00"
    
    For I = 1 To Len(Temp1)
        Temp = "&H" + (Mid$(Temp1, I, 2))
        If Left$(Temp, 3) = "&HR" Then
            Exit For
        End If
        I = I + 2
        lblSNd.Caption = Right$(Temp, 2)
        TempAtr = CInt(Temp)
        ConvertAtr
        Temp2 = Chr(HoldAtr)
        Position = Position + 1
        PBar.Value = Position
        StateChanged = 0
        If I = 264 Then
            DontReceive = True
            BufLen = 3
        Else
            BufLen = 1
        End If
        MaxP2Limit = 187
        MaxP3Limit = 177
        'Comm.RThreshold = 1
        Comm.Output = Temp2
        TimeOut = 0
        Timer1.Enabled = False
        Timer1.Interval = 75
        Timer1.Enabled = True
        Do While (TimeOut = 0) Or (StateChanged = 0)
            DoEvents
        Loop
    Next I

    Temp = ""
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
    Next I
    
    InBuf = ""
    
    'Comm.Output = ""
    lblSNd.Caption = ""
    
' well remember to add read code
'to verify write operation

    If InStr(1, Temp, "90 80") Then
        txtStat.Text = "Successful"
    Else
            txtStat.Text = "Read was unsuccessful."
            StartFlag = False
            Temp = Comm.Input
            EnableControls
            Exit Sub
    End If
    Temp = Comm.Input
    Label7.Enabled = True
    Label7.Visible = True
    EnableControls
End Sub

Sub Setupvars()
    BlockAdd(16) = "81 00 "
    BlockSig(16) = "91 65 91 99 83 "
    BlockAdd(17) = "81 10 "
    BlockSig(17) = "3B 60 FC 9A 23 "
    BlockAdd(18) = "81 20 "
    BlockSig(18) = "1F FB 3A A5 D4 "
    BlockAdd(19) = "81 30 "
    BlockSig(19) = "90 15 29 CF 97 "
    BlockAdd(20) = "81 40 "
    BlockSig(20) = "63 CC D4 19 3B "
    BlockAdd(21) = "81 50 "
    BlockSig(21) = "C3 F9 E8 97 29 "
    BlockAdd(22) = "81 60 "
    BlockSig(22) = "BC E5 09 E0 D3 "
    BlockAdd(23) = "81 70 "
    BlockSig(23) = "77 CD CB 11 C4 "
    BlockAdd(24) = "81 80 "
    BlockSig(24) = "8B A0 F0 E0 10 "
    BlockAdd(25) = "81 90 "
    BlockSig(25) = "93 A4 6F 40 48 "
    BlockAdd(26) = "81 A0 "
    BlockSig(26) = "72 58 6C 21 88 "
    BlockAdd(27) = "81 B0 "
    BlockSig(27) = "91 1C 93 92 E3 "
    BlockAdd(28) = "81 C0 "
    BlockSig(28) = "37 DB 69 5E 2E "
    BlockAdd(29) = "81 D0 "
    BlockSig(29) = "04 57 EE DD 39 "
    BlockAdd(30) = "81 E0 "
    BlockSig(30) = "12 E5 C7 F0 8F "
    BlockAdd(31) = "81 F0 "
    BlockSig(31) = "19 0C 0F 68 01 "
    
    BlockAdd(32) = "82 00 "
    BlockSig(32) = "B4 86 D4 89 E8 "
    BlockAdd(33) = "82 10 "
    BlockSig(33) = "03 F2 7E 73 86 "
    BlockAdd(34) = "82 20 "
    BlockSig(34) = "31 BC 6B D7 0D "
    BlockAdd(35) = "82 30 "
    BlockSig(35) = "F7 59 5E 6A 1E "
    BlockAdd(36) = "82 40 "
    BlockSig(36) = "A2 58 D4 F7 CE "
    BlockAdd(37) = "82 50 "
    BlockSig(37) = "E6 4E A1 14 DC "
    BlockAdd(38) = "82 60 "
    BlockSig(38) = "71 B9 38 32 8F "
    BlockAdd(39) = "82 70 "
    BlockSig(39) = "8D BB DE E6 F3 "
    BlockAdd(40) = "82 80 "
    BlockSig(40) = "2F A4 B6 4B F2 "
    BlockAdd(41) = "82 90 "
    BlockSig(41) = "C1 BD 35 DF F1 "
    BlockAdd(42) = "82 A0 "
    BlockSig(42) = "49 9C C5 AC 34 "
    BlockAdd(43) = "82 B0 "
    BlockSig(43) = "3D E8 BE E7 39 "
    BlockAdd(44) = "82 C0 "
    BlockSig(44) = "74 A8 05 FF 0F "
    BlockAdd(45) = "82 D0 "
    BlockSig(45) = "2D B2 46 2B 0F "
    BlockAdd(46) = "82 E0 "
    BlockSig(46) = "00 C2 7F 29 DF "
    
    BlockAdd(56) = "83 80 "
    BlockSig(56) = "B3 8B 97 E4 63 "
    BlockAdd(57) = "83 90 "
    BlockSig(57) = "35 25 75 9B 5C "
    BlockAdd(58) = "83 A0 "
    BlockSig(58) = "65 DE 66 E8 2B "
    BlockAdd(59) = "83 B0 "
    BlockSig(59) = "38 95 E0 32 A0 "
    BlockAdd(60) = "83 C0 "
    BlockSig(60) = "51 F5 A1 96 21 "
    BlockAdd(61) = "83 D0 "
    BlockSig(61) = "80 30 88 33 1C "
    BlockAdd(62) = "83 E0 "
    BlockSig(62) = "78 93 76 3B 78 "
    BlockAdd(63) = "83 F0 "
    BlockSig(63) = "B9 22 F3 F1 08 "
    
    BlockAdd(64) = "84 00 "
    BlockSig(64) = "57 5B EC A4 6E "
    BlockAdd(65) = "84 10 "
    BlockSig(65) = "05 D6 3E B6 22 "
    BlockAdd(66) = "84 20 "
    BlockSig(66) = "2A 90 CF 58 3F "
    BlockAdd(67) = "84 30 "
    BlockSig(67) = "3C 66 31 2E DB "
    BlockAdd(68) = "84 40 "
    BlockSig(68) = "9D AE 7D B7 45 "

    BlockAdd(80) = "85 00 "
    BlockSig(80) = "AF DF CB C2 F9 "
    BlockAdd(81) = "85 10 "
    BlockSig(81) = "86 F5 87 86 60 "
    BlockAdd(82) = "85 20 "
    BlockSig(82) = "E7 69 DF 7B B3 "
    BlockAdd(83) = "85 30 "
    BlockSig(83) = "9F 40 56 10 C3 "
    BlockAdd(84) = "85 40 "
    BlockSig(84) = "07 EE 77 49 FA "
    BlockAdd(85) = "85 50 "
    BlockSig(85) = "CB CF 0D 31 6B "
    BlockAdd(86) = "85 60 "
    BlockSig(86) = "B2 83 1E 91 7A "
    BlockAdd(87) = "85 70 "
    BlockSig(87) = "5A 95 D7 D5 8C "
    BlockAdd(88) = "85 80 "
    BlockSig(88) = "FB 97 CF 98 97 "
    BlockAdd(89) = "85 90 "
    BlockSig(89) = "DB 6C FB 63 6D "
    BlockAdd(90) = "85 A0 "
    BlockSig(90) = "E3 76 AE 63 7B "
    BlockAdd(91) = "85 B0 "
    BlockSig(91) = "20 39 C0 6F 01 "
    BlockAdd(92) = "85 C0 "
    BlockSig(92) = "06 5A C7 5C 1C "
    BlockAdd(93) = "85 D0 "
    BlockSig(93) = "B4 CE A0 E7 12 "
    BlockAdd(94) = "85 E0 "
    BlockSig(94) = "47 2E 2A FF 13 "
    BlockAdd(95) = "85 F0 "
    BlockSig(95) = "AD 16 3E 87 DF "
    
    BlockAdd(96) = "86 00 "
    BlockSig(96) = "B1 57 DC 13 21 "
    BlockAdd(97) = "86 10 "
    BlockSig(97) = "78 3D 8F CA 93 "
    BlockAdd(98) = "86 20 "
    BlockSig(98) = "1E 1E 8F DA 38 "
    BlockAdd(99) = "86 30 "
    BlockSig(99) = "44 3C A9 81 BA "
    BlockAdd(100) = "86 40 "
    BlockSig(100) = "95 49 5C 32 E6 "
    BlockAdd(101) = "86 50 "
    BlockSig(101) = "E4 37 74 D3 7E "
    BlockAdd(102) = "86 60 "
    BlockSig(102) = "71 ED 50 2D 01 "
    BlockAdd(103) = "86 70 "
    BlockSig(103) = "1A 91 66 C9 1E "
    BlockAdd(104) = "86 80 "
    BlockSig(104) = "90 3A 2C 1B 08 "
    BlockAdd(105) = "86 90 "
    BlockSig(105) = "AB 4B 9D 1E 1C "
    BlockAdd(106) = "86 A0 "
    BlockSig(106) = "E4 A5 BC A0 53 "
    BlockAdd(107) = "86 B0 "
    BlockSig(107) = "C5 C1 5E 8E 20 "
    BlockAdd(108) = "86 C0 "
    BlockSig(108) = "F2 49 AA 67 11 "
    BlockAdd(109) = "86 D0 "
    BlockSig(109) = "2E 10 B3 07 2A "
    BlockAdd(110) = "86 E0 "
    BlockSig(110) = "16 59 E4 16 2C "
    BlockAdd(111) = "86 F0 "
    BlockSig(111) = "B1 3B E5 6E 0D "
    
    BlockAdd(112) = "87 00 "
    BlockSig(112) = "F9 A4 01 67 06 "
    BlockAdd(113) = "87 10 "
    BlockSig(113) = "60 C3 DA 1E 02 "
    BlockAdd(114) = "87 20 "
    BlockSig(114) = "AF 83 A3 4B 5D "
    BlockAdd(115) = "87 30 "
    BlockSig(115) = "76 46 A2 C8 40 "
    BlockAdd(116) = "87 40 "
    BlockSig(116) = "9F 60 82 1B 38 "
    BlockAdd(117) = "87 50 "
    BlockSig(117) = "C8 E3 8E 9A 0A "
    BlockAdd(118) = "87 60 "
    BlockSig(118) = "A5 CD 76 34 38 "
    BlockAdd(119) = "87 70 "
    BlockSig(119) = "18 AD BC 93 9B "
    BlockAdd(120) = "87 80 "
    BlockSig(120) = "97 D1 2C 63 91 "
    BlockAdd(121) = "87 90 "
    BlockSig(121) = "BE E9 76 0B 14 "
    BlockAdd(122) = "87 A0 "
    BlockSig(122) = "29 99 95 A1 06 "
    BlockAdd(123) = "87 B0 "
    BlockSig(123) = "27 1A FB EE 2D "
    BlockAdd(124) = "87 C0 "
    BlockSig(124) = "6C F8 5F D1 48 "
    BlockAdd(125) = "87 D0 "
    BlockSig(125) = "EA 41 2D 83 57 "
    BlockAdd(126) = "87 E0 "
    BlockSig(126) = "BC C0 51 72 29 "
    BlockAdd(127) = "87 F0 "
    BlockSig(127) = "71 B8 58 2A B0 "
    
    BlockAdd(128) = "88 00 "
    BlockSig(128) = "D3 79 E8 7E 4B "
    BlockAdd(129) = "88 10 "
    BlockSig(129) = "28 EA 62 CC D0 "
    BlockAdd(130) = "88 20 "
    BlockSig(130) = "70 B8 F3 75 2C "
    BlockAdd(131) = "88 30 "
    BlockSig(131) = "50 86 7D 76 0B "
    BlockAdd(132) = "88 40 "
    BlockSig(132) = "B4 DE DA 57 1E "
    
    BlockAdd(149) = "89 50 "
    BlockSig(149) = "0E DC AD 55 FE "
    BlockAdd(150) = "89 60 "
    BlockSig(150) = "58 93 88 98 1D "
    BlockAdd(151) = "89 70 "
    BlockSig(151) = "51 6E 5A 54 FF "
    BlockAdd(152) = "89 80 "
    BlockSig(152) = "65 8E 71 A8 10 "
    BlockAdd(153) = "89 90 "
    BlockSig(153) = "0C B4 4D 03 EB "
    BlockAdd(154) = "89 A0 "
    BlockSig(154) = "17 3C DA 39 D7 "
    BlockAdd(155) = "89 B0 "
    BlockSig(155) = "F5 82 0C 34 88 "
    BlockAdd(156) = "89 C0 "
    BlockSig(156) = "15 BA DF 16 3A "
    BlockAdd(157) = "89 D0 "
    BlockSig(157) = "48 28 BA 51 33 "
    BlockAdd(158) = "89 E0 "
    BlockSig(158) = "E8 FE AD EA CD "
    BlockAdd(159) = "89 F0 "
    BlockSig(159) = "AA 8A 67 4B 64 "

    BlockAdd(160) = "8A 00 "
    BlockSig(160) = "A5 C8 D1 2F DD "
    BlockAdd(161) = "8A 10 "
    BlockSig(161) = "2D B2 25 B8 F4 "
    BlockAdd(162) = "8A 20 "
    BlockSig(162) = "A8 50 47 15 2F "
    BlockAdd(163) = "8A 30 "
    BlockSig(163) = "CA 70 8E DC 10 "
    BlockAdd(164) = "8A 40 "
    BlockSig(164) = "FE FA B1 7E A6 "
    BlockAdd(165) = "8A 50 "
    BlockSig(165) = "D4 7C 2F B1 3C "
    BlockAdd(166) = "8A 60 "
    BlockSig(166) = "1F 27 22 C8 99 "
    BlockAdd(167) = "8A 70 "
    BlockSig(167) = "22 41 F5 24 19 "
    BlockAdd(168) = "8A 80 "
    BlockSig(168) = "DA B2 15 36 F0 "
    BlockAdd(169) = "8A 90 "
    BlockSig(169) = "18 2F DA E4 D8 "
    BlockAdd(170) = "8A A0 "
    BlockSig(170) = "A3 7B EB FF 3E "
    BlockAdd(171) = "8A B0 "
    BlockSig(171) = "18 80 94 27 F0 "
    BlockAdd(172) = "8A C0 "
    BlockSig(172) = "59 5B 36 A6 9E "
    BlockAdd(173) = "8A D0 "
    BlockSig(173) = "05 CE 04 15 D3 "
    BlockAdd(174) = "8A E0 "
    BlockSig(174) = "96 75 58 E5 B5 "
    BlockAdd(175) = "8A F0 "
    BlockSig(175) = "C3 DF 8F 8D 48 "

    BlockAdd(176) = "8B 00 "
    BlockSig(176) = "FB D1 E0 6D 37 "
    BlockAdd(177) = "8B 10 "
    BlockSig(177) = "12 3D 36 94 D1 "
    BlockAdd(178) = "8B 20 "
    BlockSig(178) = "22 A5 19 B9 EC "
    BlockAdd(179) = "8B 30 "
    BlockSig(179) = "69 49 0A 95 6D "
    BlockAdd(180) = "8B 40 "
    BlockSig(180) = "E4 65 AF D4 39 "
    BlockAdd(181) = "8B 50 "
    BlockSig(181) = "37 93 1E CF C8 "
    BlockAdd(182) = "8B 60 "
    BlockSig(182) = "8B B0 F9 AC E2 "
    BlockAdd(183) = "8B 70 "
    BlockSig(183) = "24 C9 4A 19 1C "
    BlockAdd(184) = "8B 80 "
    BlockSig(184) = "67 1E 9D 05 80 "
    BlockAdd(185) = "8B 90 "
    BlockSig(185) = "91 01 9A 52 70 "
    BlockAdd(186) = "8B A0 "
    BlockSig(186) = "FC A6 60 31 79 "
    BlockAdd(187) = "8B B0 "
    BlockSig(187) = "70 27 76 A9 A0 "
    BlockAdd(188) = "8B C0 "
    BlockSig(188) = "D1 FC C3 D7 64 "
    BlockAdd(189) = "8B D0 "
    BlockSig(189) = "0C 06 84 35 67 "
    BlockAdd(190) = "8B E0"
    BlockSig(190) = "00 00 00 00 00"
    BlockAdd(191) = "8B F0 "
    BlockSig(191) = "4A 08 F8 13 D0 "

    BlockAdd(192) = "8C 00 "
    BlockSig(192) = "EB 40 2C 98 7B "
    BlockAdd(193) = "8C 10 "
    BlockSig(193) = "ED 9D C5 80 7B "
    BlockAdd(194) = "8C 20 "
    BlockSig(194) = "13 96 B7 58 63 "
    BlockAdd(195) = "8C 30 "
    BlockSig(195) = "99 2F 1A 85 14 "
    BlockAdd(196) = "8C 40 "
    BlockSig(196) = "97 F6 DA 99 D1 "
    BlockAdd(197) = "8C 50 "
    BlockSig(197) = "5A 2E 27 09 25 "
    BlockAdd(198) = "8C 60 "
    BlockSig(198) = "E2 47 84 78 49 "
    BlockAdd(199) = "8C 70 "
    BlockSig(199) = "BB 64 D5 BD 10 "
    BlockAdd(200) = "8C 80 "
    BlockSig(200) = "8C AA 5E 85 DB "
    BlockAdd(201) = "8C 90 "
    BlockSig(201) = "77 0F 17 6A 7A "
    BlockAdd(202) = "8C A0 "
    BlockSig(202) = "FC 78 D2 D9 5E "
    BlockAdd(203) = "8C B0 "
    BlockSig(203) = "8C C1 A2 09 1A "
    BlockAdd(204) = "8C C0 "
    BlockSig(204) = "00 AA E5 97 2F "
    BlockAdd(205) = "8C D0 "
    BlockSig(205) = "D6 03 41 38 81 "
    BlockAdd(206) = "8C E0 "
    BlockSig(206) = "F5 42 61 76 11 "
    BlockAdd(207) = "8C F0 "
    BlockSig(207) = "6B E0 8E FD 21 "

    BlockAdd(208) = "8D 00 "
    BlockSig(208) = "E9 68 8B D6 CD "
    BlockAdd(209) = "8D 10 "
    BlockSig(209) = "72 A7 92 71 CB "
    BlockAdd(210) = "8D 20 "
    BlockSig(210) = "A7 7E D5 89 8A "
    BlockAdd(211) = "8D 30 "
    BlockSig(211) = "1D 86 E8 6C AE "
    BlockAdd(212) = "8D 40 "
    BlockSig(212) = "BA B9 40 83 C4 "
    BlockAdd(213) = "8D 50 "
    BlockSig(213) = "50 C4 E7 87 66 "
    BlockAdd(214) = "8D 60 "
    BlockSig(214) = "FA 38 F9 8A 5E "
    BlockAdd(215) = "8D 70 "
    BlockSig(215) = "FE 97 A6 06 F9 "
    BlockAdd(216) = "8D 80 "
    BlockSig(216) = "7F 2C 4A 63 7F "
    BlockAdd(217) = "8D 90 "
    BlockSig(217) = "6B C1 1E 37 10 "
    BlockAdd(218) = "8D A0 "
    BlockSig(218) = "85 5C BC 8A C1 "
    BlockAdd(219) = "8D B0 "
    BlockSig(219) = "03 20 60 F5 76 "
    BlockAdd(220) = "8D C0 "
    BlockSig(220) = "E3 B9 84 B1 1B "
    BlockAdd(221) = "8D D0 "
    BlockSig(221) = "F4 FA C9 32 39 "
    BlockAdd(222) = "8D E0 "
    BlockSig(222) = "D7 40 DC D3 52 "
    BlockAdd(223) = "8D F0 "
    BlockSig(223) = "6D 16 61 88 6F "

    BlockAdd(224) = "8E 00 "
    BlockSig(224) = "2F 10 1B 68 4F "
    BlockAdd(225) = "8E 10 "
    BlockSig(225) = "CE 10 FE D9 22 "
    BlockAdd(226) = "8E 20 "
    BlockSig(226) = "90 33 82 2B D0 "
    BlockAdd(227) = "8E 30 "
    BlockSig(227) = "14 8C 65 6B 6E "
    BlockAdd(228) = "8E 40 "
    BlockSig(228) = "C4 65 8B CC 32 "
    BlockAdd(229) = "8E 50 "
    BlockSig(229) = "2C A3 69 74 1B "
    BlockAdd(230) = "8E 60 "
    BlockSig(230) = "47 18 BC 92 CC "
    BlockAdd(231) = "8E 70 "
    BlockSig(231) = "7F 40 05 3B 54 "
    BlockAdd(232) = "8E 80 "
    BlockSig(232) = "3D CE FB D8 F8 "
    BlockAdd(233) = "8E 90 "
    BlockSig(233) = "65 70 D9 15 E9 "
    BlockAdd(234) = "8E A0 "
    BlockSig(234) = "65 5A 88 6D 39 "
    BlockAdd(235) = "8E B0 "
    BlockSig(235) = "9E 98 AC 20 C3 "
    BlockAdd(236) = "8E C0 "
    BlockSig(236) = "A5 89 CC 41 6B "
    BlockAdd(237) = "8E D0 "
    BlockSig(237) = "E6 BD 51 6D 84 "
    BlockAdd(238) = "8E E0 "
    BlockSig(238) = "ED F7 81 5F 55 "
    BlockAdd(239) = "8E F0 "
    BlockSig(239) = "CF 65 06 07 F7 "

    BlockAdd(240) = "8F 00 "
    BlockSig(240) = "DA 4D 2C 96 5F "
    BlockAdd(241) = "8F 10 "
    BlockSig(241) = "3E 40 D3 A4 20 "
    BlockAdd(242) = "8F 20 "
    BlockSig(242) = "B2 00 8A 25 0A"
    BlockAdd(243) = "8F 30 "
    BlockSig(243) = "21 6D DB 45 08 "
    BlockAdd(244) = "8F 40 "
    BlockSig(244) = "51 78 19 55 42 "
    BlockAdd(245) = "8F 50 "
    BlockSig(245) = "2A DE 01 F3 00 "
    BlockAdd(246) = "8F 60 "
    BlockSig(246) = "16 61 07 D7 5A "
    BlockAdd(247) = "8F 70 "
    BlockSig(247) = "04 E7 0E 56 47 "
    BlockAdd(248) = "8F 80 "
    BlockSig(248) = "6C 63 45 B1 E3 "
    BlockAdd(249) = "8F 90 "
    BlockSig(249) = "9D 18 EC 3B 66 "
    BlockAdd(250) = "8F A0 "
    BlockSig(250) = "F5 F2 BD 42 A1 "
    BlockAdd(251) = "8F B0 "
    BlockSig(251) = "4D D4 C5 F0 55 "
    BlockAdd(252) = "8F C0 "
    BlockSig(252) = "DA 9D 97 39 3A "
    BlockAdd(253) = "8F D0 "
    BlockSig(253) = "99 21 A9 80 4E "
    BlockAdd(254) = "8F E0 "
    BlockSig(254) = "15 13 24 12 F4 "
    BlockAdd(255) = "8F F0 "
    BlockSig(255) = "6C 16 91 4B CF "

    Un29 = "48 40 00 00 C2" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "R01" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "09 10 00 00 30 60 C0 B2 00 1E 85 08 0B 85 6C 7F" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "0C 12 1C C1 79 80 02 86 13 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "00 00 00 00 00 00 00 00 00 00 BB 00 0C BD 4E E4" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "C6 63" + Chr$(13) + Chr$(10)
    Un29 = Un29 + "R02"
    
    TestE3 = "48 42 00 00  08" + Chr$(13) + Chr$(10)
    TestE3 = TestE3 + "R01" + Chr$(13) + Chr$(10)
    TestE3 = TestE3 + "30 60 D5 02 07 2C E3 00" + Chr$(13) + Chr$(10)
    TestE3 = TestE3 + "R02" + Chr$(13) + Chr$(10)
    TestE3 = TestE3 + "R09"
    
    RdE3 = "48 40 00 00 58" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "R01" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 " + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "bb 06 01 85 91 80 00 00" + Chr$(13) + Chr$(10)
    RdE3 = RdE3 + "R02" + Chr$(13) + Chr$(10)
    
    RdE31 = RdE31 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "R01" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "60 D5 02 85 8E E3 12 0E 80 F0 90 80 00 7F 01 12 07 8A E5 83 B4 82 F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "R02" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "R01" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "R02" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "RFF" + Chr$(13) + Chr$(10)
    RdE31 = RdE31 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE32 = RdE32 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "R01" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "60 D5 02 85 8E E3 12 0E 80 F0 90 82 00 7F 01 12 07 8A E5 83 B4 84 F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "R02" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "R01" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "R02" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "RFF" + Chr$(13) + Chr$(10)
    RdE32 = RdE32 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE33 = RdE33 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "R01" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "60 D5 02 85 8E E3 12 0E 80 F0 90 84 00 7F 01 12 07 8A E5 83 B4 86 F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "R02" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "R01" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "R02" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "RFF" + Chr$(13) + Chr$(10)
    RdE33 = RdE33 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE34 = RdE34 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "R01" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "60 D5 02 85 8E E3 12 0E 80 F0 90 86 00 7F 01 12 07 8A E5 83 B4 88 F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "R02" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "R01" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "R02" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "RFF" + Chr$(13) + Chr$(10)
    RdE34 = RdE34 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE35 = RdE35 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "R01" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "60 D5 02 85 8E E3 12 0E 80 F0 90 88 00 7F 01 12 07 8A E5 83 B4 8A F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "R02" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "R01" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "R02" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "RFF" + Chr$(13) + Chr$(10)
    RdE35 = RdE35 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE36 = RdE36 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "R01" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "60 D5 02 85 8E E3 12 0E 80 F0 90 8A 00 7F 01 12 07 8A E5 83 B4 8C F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "R02" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "R01" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "R02" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "RFF" + Chr$(13) + Chr$(10)
    RdE36 = RdE36 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE37 = RdE37 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "R01" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "60 D5 02 85 8E E3 12 0E 80 F0 90 8C 00 7F 01 12 07 8A E5 83 B4 8E F6 22 00  BB 00" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "R02" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "R01" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "R02" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "RFF" + Chr$(13) + Chr$(10)
    RdE37 = RdE37 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE38 = RdE38 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "R01" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "60 D5 02 85 8E E3 12 0E 80 F0 90 8E 00 7F 01 12 07 8A E5 83 B4 90 F6 22 00 BB 00" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "R02" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "48 40 00 00 07" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "R01" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "60 D5 02 80 6B E3 00" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "R02" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "RFF" + Chr$(13) + Chr$(10)
    RdE38 = RdE38 + "RFF" + Chr$(13) + Chr$(10)
    
    RdE39 = RdE39 + "48 42 00 00 1B" + Chr$(13) + Chr$(10)
    RdE39 = RdE39 + "R01" + Chr$(13) + Chr$(10)
    RdE39 = RdE39 + "60 D5 02 85 8E E3 12 0E 80 F0 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 BB 00" + Chr$(13) + Chr$(10)
    RdE39 = RdE39 + "R02" + Chr$(13) + Chr$(10)
    
    Chk29 = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    Chk29 = Chk29 + "R01" + Chr$(13) + Chr$(10)
    Chk29 = Chk29 + "09 10 00 00 24 25 60 B5 03 85 C8 02 BB 00 0C 1E" + Chr$(13) + Chr$(10)
    Chk29 = Chk29 + "D6 A5 7E 2F" + Chr$(13) + Chr$(10)
    Chk29 = Chk29 + "R02" + Chr$(13) + Chr$(10)
    
    SafeCheck(1) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(1) = SafeCheck(1) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(1) = SafeCheck(1) + "09 10 00 00 24 25 60 B5 03 80 10 1E BB 00 0C FD" + Chr$(13) + Chr$(10)
    SafeCheck(1) = SafeCheck(1) + "F0 29 FB A5" + Chr$(13) + Chr$(10)
    SafeCheck(1) = SafeCheck(1) + "R02"

    SafeCheck(2) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(2) = SafeCheck(2) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(2) = SafeCheck(2) + "09 10 00 00 24 25 60 B5 03 81 00 20 BB 00 0C 28" + Chr$(13) + Chr$(10)
    SafeCheck(2) = SafeCheck(2) + "2A 9B 3B DF" + Chr$(13) + Chr$(10)
    SafeCheck(2) = SafeCheck(2) + "R02"

    SafeCheck(3) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(3) = SafeCheck(3) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(3) = SafeCheck(3) + "09 10 00 00 24 25 60 B5 03 82 00 1E BB 00 0C 18" + Chr$(13) + Chr$(10)
    SafeCheck(3) = SafeCheck(3) + "C5 2A D7 8E" + Chr$(13) + Chr$(10)
    SafeCheck(3) = SafeCheck(3) + "R02"

    SafeCheck(4) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(4) = SafeCheck(4) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(4) = SafeCheck(4) + "09 10 00 00 24 25 60 B5 03 83 80 10 BB 00 0C F6" + Chr$(13) + Chr$(10)
    SafeCheck(4) = SafeCheck(4) + "6A 70 4D 45" + Chr$(13) + Chr$(10)
    SafeCheck(4) = SafeCheck(4) + "R02"

    SafeCheck(5) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(5) = SafeCheck(5) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(5) = SafeCheck(5) + "09 10 00 00 24 25 60 B5 03 84 00 0A BB 00 0C 93" + Chr$(13) + Chr$(10)
    SafeCheck(5) = SafeCheck(5) + "AA 72 FF 56" + Chr$(13) + Chr$(10)
    SafeCheck(5) = SafeCheck(5) + "R02"

    SafeCheck(6) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(6) = SafeCheck(6) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(6) = SafeCheck(6) + "09 10 00 00 24 25 60 B5 03 85 00 20 BB 00 0C C6" + Chr$(13) + Chr$(10)
    SafeCheck(6) = SafeCheck(6) + "09 89 3B EC" + Chr$(13) + Chr$(10)
    SafeCheck(6) = SafeCheck(6) + "R02"

    SafeCheck(7) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(7) = SafeCheck(7) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(7) = SafeCheck(7) + "09 10 00 00 24 25 60 B5 03 86 00 20 BB 00 0C AC" + Chr$(13) + Chr$(10)
    SafeCheck(7) = SafeCheck(7) + "B8 FC 69 2E" + Chr$(13) + Chr$(10)
    SafeCheck(7) = SafeCheck(7) + "R02"

    SafeCheck(8) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(8) = SafeCheck(8) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(8) = SafeCheck(8) + "09 10 00 00 24 25 60 B5 03 87 00 20 BB 00 0C 66" + Chr$(13) + Chr$(10)
    SafeCheck(8) = SafeCheck(8) + "02 80 6F 40" + Chr$(13) + Chr$(10)
    SafeCheck(8) = SafeCheck(8) + "R02"

    SafeCheck(9) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(9) = SafeCheck(9) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(9) = SafeCheck(9) + "09 10 00 00 24 25 60 B5 03 88 00 0A BB 00 0C CC" + Chr$(13) + Chr$(10)
    SafeCheck(9) = SafeCheck(9) + "94 07 F8 2D" + Chr$(13) + Chr$(10)
    SafeCheck(9) = SafeCheck(9) + "R02"
    
    SafeCheck(10) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(10) = SafeCheck(10) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(10) = SafeCheck(10) + "09 10 00 00 24 25 60 B5 03 89 50 16 BB 00 0C 12" + Chr$(13) + Chr$(10)
    SafeCheck(10) = SafeCheck(10) + "F0 78 60 A8" + Chr$(13) + Chr$(10)
    SafeCheck(10) = SafeCheck(10) + "R02"

    SafeCheck(11) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(11) = SafeCheck(11) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(11) = SafeCheck(11) + "09 10 00 00 24 25 60 B5 03 8A 00 20 BB 00 0C 6C" + Chr$(13) + Chr$(10)
    SafeCheck(11) = SafeCheck(11) + "F0 F2 E9 94" + Chr$(13) + Chr$(10)
    SafeCheck(11) = SafeCheck(11) + "R02"

    SafeCheck(12) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(12) = SafeCheck(12) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(12) = SafeCheck(12) + "09 10 00 00 24 25 60 B5 03 8B 00 20 BB 00 0C E6" + Chr$(13) + Chr$(10)
    SafeCheck(12) = SafeCheck(12) + "97 DC 08 C3" + Chr$(13) + Chr$(10)
    SafeCheck(12) = SafeCheck(12) + "R02"

    SafeCheck(13) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(13) = SafeCheck(13) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(13) = SafeCheck(13) + "09 10 00 00 24 25 60 B5 03 8C 00 20 BB 00 0C 2A" + Chr$(13) + Chr$(10)
    SafeCheck(13) = SafeCheck(13) + "3E ED 7A 46" + Chr$(13) + Chr$(10)
    SafeCheck(13) = SafeCheck(13) + "R02"

    SafeCheck(14) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(14) = SafeCheck(14) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(14) = SafeCheck(14) + "09 10 00 00 24 25 60 B5 03 8D 00 20 BB 00 0C 66" + Chr$(13) + Chr$(10)
    SafeCheck(14) = SafeCheck(14) + "8E 7B D8 64" + Chr$(13) + Chr$(10)
    SafeCheck(14) = SafeCheck(14) + "R02"

    SafeCheck(15) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(15) = SafeCheck(15) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(15) = SafeCheck(15) + "09 10 00 00 24 25 60 B5 03 8E 00 20 BB 00 0C DA" + Chr$(13) + Chr$(10)
    SafeCheck(15) = SafeCheck(15) + "9B 55 25 EE" + Chr$(13) + Chr$(10)
    SafeCheck(15) = SafeCheck(15) + "R02"

    SafeCheck(16) = "48 42 00 00 14" + Chr$(13) + Chr$(10)
    SafeCheck(16) = SafeCheck(16) + "R01" + Chr$(13) + Chr$(10)
    SafeCheck(16) = SafeCheck(16) + "09 10 00 00 24 25 60 B5 03 8F 00 20 BB 00 0C CA" + Chr$(13) + Chr$(10)
    SafeCheck(16) = SafeCheck(16) + "D6 2A B6 83" + Chr$(13) + Chr$(10)
    SafeCheck(16) = SafeCheck(16) + "R02"

    Janitor(1) = "48 40 00 00 67" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "R01" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "BB 15 10 80 28 74 37 12 05 F7 12 06 8A B4 A5 F7" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "12 06 8A C0 E0 00 00" + Chr$(13) + Chr$(10)
    Janitor(1) = Janitor(1) + "R02"
    
    Janitor(2) = "48 40 00 00 67" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "R01" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "BB 15 10 80 38 54 F0 B4 80 1B 12 06 8A C0 E0 12" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "06 8A F5 0F D0 00 00" + Chr$(13) + Chr$(10)
    Janitor(2) = Janitor(2) + "R02"
    
    Janitor(3) = "48 40 00 00 67" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "R01" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "BB 15 10 80 48 E0 F5 82 D0 E0 F5 83 79 0F 12 16" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "77 E5 0F 80 D2 00 00" + Chr$(13) + Chr(10)
    Janitor(3) = Janitor(3) + "R02" + Chr$(13) + Chr(10)
    
    Janitor(4) = "48 40 00 00 67" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "R01" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "BB 15 10 80 58 74 26 12 05 F7 74 00 C0 E0 C0 E0" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "90 80 28 7F BF 00 00" + Chr$(13) + Chr$(10)
    Janitor(4) = Janitor(4) + "R02" + Chr$(13) + Chr$(10)
    
    Janitor(5) = "48 40 00 00 5A" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "R01" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "BB 08 03 80 68 02 05 73 00 00" + Chr$(13) + Chr$(10)
    Janitor(5) = Janitor(5) + "R02" + Chr$(13) + Chr$(10)
    
    Janitor(6) = "48 40 00 00 5A" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "R01" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "BB 08 03 85 BD 02 80 28 00 00" + Chr$(13) + Chr$(10)
    Janitor(6) = Janitor(6) + "R02"
    
    Janitor(7) = "48 40 00 00 04" + Chr$(13) + Chr$(10)
    Janitor(7) = Janitor(7) + "R01" + Chr$(13) + Chr$(10)
    Janitor(7) = Janitor(7) + "30 60 BD 00" + Chr$(13) + Chr$(10)
    Janitor(7) = Janitor(7) + "R01" + Chr$(13) + Chr$(10)
    
    ' Restore the 'BD nano to its original shape.
    Janitor(8) = "A5 85 BD 02" + Chr$(13) + Chr$(10)
    Janitor(8) = Janitor(8) + "R01" + Chr$(13) + Chr$(10)
    Janitor(8) = Janitor(8) + "A5 85 BE 85" + Chr$(13) + Chr$(10)
    Janitor(8) = Janitor(8) + "R01" + Chr$(13) + Chr$(10)
    Janitor(8) = Janitor(8) + "A5 85 BF 60" + Chr$(13) + Chr$(10)
    Janitor(8) = Janitor(8) + "R01" + Chr$(13) + Chr$(10)
        
End Sub

Sub CheckOpen()
    Dim I As Integer
    Dim Temp As String
    Dim pos As Integer
    Dim pos2 As Integer

    InBuf = ""
    EEProm = ""
    XPLList.Clear
    txtStat.Text = "Closing open holes..."
    ' close e3
    Temp = "48 42 00 00 58"
    XPLList.AddItem Temp
    Temp = "R01"
    XPLList.AddItem Temp
    Temp = "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03"
    XPLList.AddItem Temp
    Temp = " 0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e"
    XPLList.AddItem Temp
    Temp = " 1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 "
    XPLList.AddItem Temp
    Temp = "bb 06 01 85 91 90 00 00"
    XPLList.AddItem Temp
    Temp = "R02"
    XPLList.AddItem Temp
    
    SendEEprom = True
    RunXPL
        
    If InStr(1, InBuf, "90 80 ") = False Then
        GoTo SendEnd
    End If
    
    'close 09 hole
    XPLList.Clear
    Temp = "48 42 00 00 58"
    XPLList.AddItem Temp
    Temp = "R01"
    XPLList.AddItem Temp
    Temp = "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03"
    XPLList.AddItem Temp
    Temp = " 0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e"
    XPLList.AddItem Temp
    Temp = " 1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 "
    XPLList.AddItem Temp
    Temp = "bb 06 01 8F 2F 09 00 00"
    XPLList.AddItem Temp
    Temp = "R02"
    XPLList.AddItem Temp
    
    SendEEprom = True
    RunXPL
        
    If InStr(1, InBuf, "90 80 ") = False Then
        GoTo SendEnd
    End If
    Label8.Enabled = False
    txtStat.Text = "Holes closed successfully."
    Exit Sub
    
SendEnd:
    txtStat.Text = "Unable to close all holes."
End Sub

Sub SafetyCheck()
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim I As Integer
    Dim J As Long
    Dim Count As Integer
    Dim Block(16) As Integer
    Dim BadBlocks(256) As Integer
    Dim cnt As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    Dim Hold As String
    Dim Elist As String
    Dim Hold1 As Integer
    Dim Hold2 As Integer
    Dim Total As Integer
    
    SendEEprom = True
    DisableControls
    
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    InBuf = ""
    EEProm = ""
    XPLList.Clear
    cnt = 0
    'XPLList.Visible = True
    'sendList.Visible = True
    For J = 1 To 256
        BadBlocks(J) = 0
    Next
        
    For J = 1 To 16
        Block(J) = 0
    Next
    
    For Count = 1 To 16
    
    XPLList.Clear
    pos = 1
    For I = 1 To 5
        pos2 = InStr(pos, SafeCheck(Count), Chr$(13) + Chr$(10))
        If I = 5 Then
            Temp = Mid$(SafeCheck(Count), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(SafeCheck(Count), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    txtStat.Text = "EEProm integrity check Block: " + Str(Count)
    RunXPL
    If Right$(InBuf, 6) <> "90 80 " Then
        cnt = cnt + 1
        Block(Count) = 1
    End If
    Next Count
    
    If cnt > 0 Then
    'don't call Janitor till we have all bad spots
        For I = 1 To 16
        XPLList.Clear
        If Block(I) = 1 Then
            txtStat.Text = "EEProm integrity check Block: " + Str$(I)
            Select Case I
                Case 2 '8100
                    For J = 16 To 31
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 3 ' 8200
                    For J = 32 To 46
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 4  ' 8300
                    For J = 56 To 63
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 5  ' 8400
                    For J = 64 To 68
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 6
                    For J = 80 To 95
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 7 ' 8600
                    For J = 96 To 111
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 8 '8700
                    For J = 112 To 127
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 9 '8800
                    For J = 128 To 132
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 10 '8900
                    For J = 149 To 159
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 11 ' 8A00
                    For J = 160 To 175
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 12  '8B00
                    For J = 176 To 191
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 13  ' 8C00
                    For J = 192 To 207
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 14  '8D00
                    For J = 208 To 223
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 15  '8E00
                    For J = 224 To 239
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
                Case 16 '8F00
                    For J = 240 To 255
                        XPLList.Clear
                        Temp = BlockHeader1
                        XPLList.AddItem Temp
                        Temp = BlockFooter1
                        XPLList.AddItem Temp
                        Temp = BlockHeader2
                        Temp = Temp + BlockAdd(J)
                        XPLList.AddItem Temp
                        Temp = BlockHeader3
                        Temp = Temp + BlockSig(J)
                        XPLList.AddItem Temp
                        XPLList.AddItem BlockFooter2
                        ResetATR
                        RunXPL
                        'here we have to track
                        'the bad blocks
                        If Right$(InBuf, 6) <> "90 80 " Then
                            'ok bad spot
                            BadBlocks(J) = 1
                        End If
                    Next J
            End Select
        End If
        Next I
        Total = 0
        Call JanitorAdd
        If JanitorFlag = False Then
            txtStat.Text = "Failed..cleaning card."
            GoTo SendEnd
        End If
        XPLList.Clear
        For J = 0 To 255
            If (BadBlocks(J) = 1) And (J <> 190) Then
                cnt = 0
                Form2.BinList.ListIndex = J
                Elist = Form2.BinList.Text
                Temp = BlockAdd(J)
                Hold = Left$(Temp, 2) + Mid$(BlockAdd(J), 4, 2)
                txtStat.Text = "Attempting to repair " + Hold
                Hold2 = ConvertHex(Right$(Hold, 2))
                Hold2 = Hold2
                For I = 1 To 48 Step 3
                    Temp = ""
                    Temp = "A5 " + Left$(Hold, 2) + " "
                    Temp = Temp + Hex(Hold2 + cnt) + " "
                    If Mid$(Elist, I, 2) <> "-1" Then
                        Temp = Temp + Mid$(Elist, I, 2)
                        XPLList.AddItem Temp
                        Temp = "R01"
                        XPLList.AddItem Temp
                        cnt = cnt + 1
                        Total = Total + cnt
                    End If
                Next I
                If cnt > 0 Then
                    RunXPL
                    XPLList.Clear
                End If
            End If
        Next J
        XPLList.Clear
        For I = &H6B To &HFF
            Temp = "A5 " + "80 " + Hex(I) + " " + "00"
            XPLList.AddItem Temp
            Temp = "R01"
            XPLList.AddItem Temp
        Next
        RunXPL
        XPLList.Clear
        Temp = "A5 FF"
        XPLList.AddItem Temp
        Temp = "R01"
        XPLList.AddItem Temp
        RunXPL
        txtStat.Text = "Repair EEProm complete. " + Str$(Total) + " bytes fixed."
    End If
    
SendEnd:
    
End Sub

Sub JanitorAdd()
    Dim I As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    Dim Index As Integer
    Dim J As Integer
    Dim Temp As String
    
    'remember present position
    J = XPLList.ListIndex
    sendList.Text = ""
    Index = XPLList.ListCount
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, Janitor(1), Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(Janitor(1), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(1), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    ResetATR
    RunXPL
    
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, Janitor(2), Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(Janitor(2), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(2), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    ResetATR
    RunXPL
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, Janitor(3), Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(Janitor(3), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(3), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    ResetATR
    RunXPL
    
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
    XPLList.Clear
    pos = 1
    For I = 1 To 10
        pos2 = InStr(pos, Janitor(4), Chr$(13) + Chr$(10))
        If I = 10 Then
            Temp = Mid$(Janitor(4), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(4), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    ResetATR
    RunXPL
    
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
    XPLList.Clear
    pos = 1
    For I = 1 To 9
        pos2 = InStr(pos, Janitor(5), Chr$(13) + Chr$(10))
        If I = 9 Then
            Temp = Mid$(Janitor(5), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(5), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    ResetATR
    RunXPL
    
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If

    XPLList.Clear
    pos = 1
    For I = 1 To 9
        pos2 = InStr(pos, Janitor(6), Chr$(13) + Chr$(10))
        If I = 9 Then
            Temp = Mid$(Janitor(6), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(6), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    sendList.Text = ""
    ResetATR
    RunXPL
    If Right$(InBuf, 6) <> "90 80 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
    XPLList.Clear
    pos = 1
    For I = 1 To 4
        pos2 = InStr(pos, Janitor(7), Chr$(13) + Chr$(10))
        If I = 4 Then
            Temp = Mid$(Janitor(7), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(7), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    sendList.Text = ""
    ResetATR
    RunXPL
    
    If Right$(InBuf, 3) <> "37 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
    XPLList.Clear
    pos = 1
    For I = 1 To 6
        pos2 = InStr(pos, Janitor(8), Chr$(13) + Chr$(10))
        If I = 6 Then
            Temp = Mid$(Janitor(8), pos, 3)
        ElseIf pos2 = 0 Then
            Exit For
        Else
            Temp = Mid$(Janitor(8), pos, pos2 - pos)
        End If
        XPLList.AddItem Temp
        pos = pos2 + 2
    Next I
    sendList.Text = ""
    RunXPL
    If Right$(InBuf, 3) <> "60 " Then
        JanitorFlag = False
        Exit Sub
    Else
        JanitorFlag = True
    End If
    
End Sub

Sub Fix745()
    Dim I As Integer
    Dim Temp As String
    Dim pos As Integer
    Dim pos2 As Integer

    InBuf = ""
    EEProm = ""
    XPLList.Clear
    'IRD Location 84FC - 84FF
    Temp = "48 42 00 00 58"
    XPLList.AddItem Temp
    Temp = "R01"
    XPLList.AddItem Temp
    Temp = "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03"
    XPLList.AddItem Temp
    Temp = " 0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e"
    XPLList.AddItem Temp
    Temp = " 1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 "
    XPLList.AddItem Temp
    Temp = "bb 0A 05 84 FC 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = "R02"
    XPLList.AddItem Temp
    
    SendEEprom = True
    ResetATR
    RunXPL
        
    If InStr(1, InBuf, "22 33 ") = False Then
        GoTo SendEnd
    End If
    'IRD Location 83D0 - 83D3
    XPLList.Clear
    'IRD Location 84FC - 84FF
    Temp = "48 42 00 00 58"
    XPLList.AddItem Temp
    Temp = "R01"
    XPLList.AddItem Temp
    Temp = "09 11 00 00 30 60 00 06 39 00 04 f4 22 33 cf 03"
    XPLList.AddItem Temp
    Temp = " 0e 1b 00 cf 03 0e 1b 00 cf 03 0e 1b 00 cf 03 0e"
    XPLList.AddItem Temp
    Temp = " 1b 00 cf 03 0e 1b 00 bb 00 12 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
    XPLList.AddItem Temp
    Temp = " 00 00 35 08 13 86 46 13 8a 1c 00 00 00 00 00 60 "
    XPLList.AddItem Temp
    Temp = "bb 0A 05 83 D0 00 00 00 00 00 00"
    XPLList.AddItem Temp
    Temp = "R02"
    XPLList.AddItem Temp
    
    SendEEprom = True
    ResetATR
    RunXPL
        
    If InStr(1, InBuf, "22 33 ") = False Then
        GoTo SendEnd
    End If
    XPLList.Clear
    Call CheckOpen
    txtStat.Text = "745 Message fixed."
    Exit Sub
    
SendEnd:
    txtStat.Text = "Unable to fix 745 Message."
End Sub
