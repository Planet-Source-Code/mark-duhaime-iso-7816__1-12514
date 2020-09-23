VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ISO7816 © 2000"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   ControlBox      =   0   'False
   FillColor       =   &H00808000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   4200
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&XIT"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   0
      ScaleHeight     =   3360
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00800000&
         Height          =   3375
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   6495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6480
      Top             =   0
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim Lines(50) As String
Dim Timeout1 As Integer
Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim TempString As String

Private Sub Form_Activate()
    Dim I As Integer
    Dim Temp As String
    Dim Start As Integer
    
    If Form4Start = True Then
        Call EEProm
        Exit Sub
    End If
        
    Timeout1 = 0
    'Command1.Visible = True
    P1.Visible = True
    
    Me.Caption = "About " + MagiString + " © 2000"
    Lines(1) = MagiString + " © 2000"
    Lines(2) = ""
    Lines(3) = "The ultimate utility for ISO7816 SmartCards"
    Lines(4) = ""
    Lines(5) = "...Repair EEProm"
    Lines(6) = "...Read and write to smartcard"
    Lines(7) = "...Run XPL files"
    Lines(8) = ""
    Lines(9) = ""
    Lines(10) = "...And much, much more."
    Lines(11) = ""
    Lines(12) = ""
    Lines(13) = ""
    Lines(14) = ""
    Lines(15) = ""
    Lines(16) = ""
    Lines(17) = "Warning and Welcome "
    Lines(18) = ""
    Lines(19) = ""
    Lines(20) = "There's a few things we would like to "
    Lines(21) = "inform you about. "
    Lines(22) = "This program is free."
    Lines(23) = "If you paid for this program then"
    Lines(24) = "you have been Scammed/RippedOff"
    Lines(25) = "/Taken. "
    Start = 1
    P1.AutoRedraw = True
    P1.Visible = False
    P1.FontSize = 12
    P1.ForeColor = &HFF0000
    P1.BackColor = BackColor
    P1.ScaleMode = 3
    ScaleMode = 3
    TempString = 49
    P1.Height = (Val(TempString) * P1.TextHeight("Test Height")) + 200
        For I = 1 To 25
        TempString = Lines(I)
        PrintText TempString
    Next I
Top:

    theleft = 0
    thetop = ScaleHeight
    p1hgt = P1.ScaleHeight
    p1wid = P1.ScaleWidth
    Timer1.Enabled = True
    Timer1.Interval = 10

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Timer1.Enabled = False
        Unload Me
    End If
End Sub

Sub Form_Load()
    Me.Refresh
End Sub

Sub Timer1_Timer()
    x% = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
    thetop = thetop - 1
    If thetop < -p1hgt Then
        Timer1.Enabled = False
        'theleft = 0
        'thetop = ScaleHeight
        'p1hgt = P1.ScaleHeight
        'p1wid = P1.ScaleWidth
        'Timer1.Enabled = True
        'Timer1.Interval = 10
        Timer1.Enabled = False
        Unload Me
    End If
End Sub

Sub PrintText(Text As String)
P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
P1.ForeColor = 0: x = P1.CurrentX: y = P1.CurrentY
For I = 1 To 3
    P1.Print Text
    x = x + 1: y = y + 1: P1.CurrentX = x: P1.CurrentY = y
Next I
P1.ForeColor = &HC0C0&
P1.Print Text
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = False
    Command1.Visible = False
    Text1.Text = ""
    Text1.Visible = False
    Timer2.Enabled = False
    Timeout2 = 1
    TheEnd = False
    Form4Start = False
    Unload Me
End Sub

Private Sub EEProm()
    Me.Caption = MagiString + " © 2000 --- Important EEProm Locations"
    Lines(1) = "8020 -- Fuse Byte"
    Lines(2) = "8028 thru 80EF  -- PPV Entries"
    Lines(3) = "80F0 thru 822F -- Tier Areas"
    Lines(4) = "82EC thru 831B -- Keys"
    Lines(5) = "8370 -- Rom Version"
    Lines(6) = "8374 thru 8377 Card Number"
    Lines(7) = "83B1 -- Ratings limit"
    Lines(8) = "83B2 -- Spending Limit"
    Lines(9) = "83B4 thru 83B7 Activation info"
    Lines(10) = "83C4 thru 83C7 Password"
    Lines(11) = "83C8 thru 83CF Zip Code"
    Lines(12) = "83D0 thru 83D3 IRD Number"
    Lines(13) = "83D4 thru 83DF PPV Slots Info"
    Lines(14) = "8406 thru 8407 Update Status Word"
    Lines(15) = "8415 TimeZone"
    Lines(16) = "8440 thru 8451 Channel Blackout Bits"
    Lines(17) = "845F Guide Byte"
    Lines(18) = "84FC thru 84FF Current IRD Number"
    Text1.Visible = True
    Command1.Visible = True
    
    For I = 1 To 18
        Text1.Text = Text1.Text + Lines(I) + Chr$(13) + Chr$(10)
    Next
Top:
    Timer1.Enabled = False

    Timeout1 = 0
    Timer2.Enabled = False
    Timer2.Interval = 60
    Do
        DoEvents
    Loop While Timeout1 = 0
    
    If TheEnd = True Then
        Exit Sub
    End If
    GoTo Top
            
End Sub

Private Sub Timer2_Timer()
    Timeout1 = 1
    Timer2.Enabled = False
End Sub

