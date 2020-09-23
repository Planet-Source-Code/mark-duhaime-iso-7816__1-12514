VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Select Bin Image to Load"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Bin image to load:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    BinFile = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()

    I = FileList.ListIndex
    BinFile = FileList.FileName
    Unload Me
    
End Sub

Private Sub FileList_DblClick()
    
    I = FileList.ListIndex
    BinFile = FileList.FileName
    Unload Me
    
End Sub

Private Sub Form_Load()
    Me.Refresh
    FileList.FileName = FilName
    FileList.Refresh
    
End Sub

