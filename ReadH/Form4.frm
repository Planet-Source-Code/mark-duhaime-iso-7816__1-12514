VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4545
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   4515
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   4440
         TabIndex        =   2
         Top             =   2400
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   425
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheEnd As Boolean

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim TempString As String

Private Sub Command1_Click()
    Timer1.Enabled = False
    Timeout1 = 1
    TheEnd = True
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    TheEnd = False
    
    If Form4Start = False Then
        Exit Sub
    End If
    
    'Me.Show
    Me.Refresh

Top:
    P1.AutoRedraw = True
    P1.FontSize = 8
    'P1.ForeColor = &HFF0000
    P1.BackColor = &H808000
    P1.ScaleMode = 3
    ScaleMode = 3
    P1.Height = 2055 '(39 * P1.TextHeight("Test Height")) + 200
    'For I = 1 To 39
    '    Tempstring = Lines(I)
    '    PrintText Tempstring
    'Next I
    
    theleft = 0
    thetop = ScaleHeight
    p1hgt = P1.ScaleHeight
    p1wid = P1.ScaleWidth
    Timer1.Enabled = True
    Timer1.Interval = 10
    
   
End Sub

Sub Timer1_Timer()
    x% = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
    thetop = thetop - 1
    If thetop < -p1hgt Then
        Timer1.Enabled = False
        Txt$ = "Credits Completed"
        CurrentY = ScaleHeight / 2
        CurrentX = (ScaleWidth - TextWidth(Txt$)) / 2
        Print Txt$
    End If
End Sub

Sub PrintText(Text As String)
    P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
    x = P1.CurrentX
    y = P1.CurrentY
    P1.ForeColor = 0
    
    For I = 1 To 3
        P1.Print Text
        x = x + 1: y = y + 1: P1.CurrentX = x: P1.CurrentY = y
    Next I
    P1.ForeColor = &HFFFF&
    P1.Print Text
End Sub

