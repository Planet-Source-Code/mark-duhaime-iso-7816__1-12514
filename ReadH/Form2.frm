VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "EEProm"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7050
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7200
      Top             =   6600
   End
   Begin VB.TextBox txtStat 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   335
      Left            =   2520
      TabIndex        =   299
      Top             =   6600
      Width           =   3135
   End
   Begin VB.ListBox UpdateList 
      Height          =   255
      Left            =   6000
      TabIndex        =   297
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCorrect 
      BackColor       =   &H00C0C0C0&
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
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   296
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   295
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox BinList 
      Height          =   255
      Left            =   6480
      TabIndex        =   293
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   10
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   0
      Top             =   10
      Width           =   10
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   255
      Left            =   6120
      TabIndex        =   288
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   254
      Left            =   5760
      TabIndex        =   287
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   253
      Left            =   5400
      TabIndex        =   286
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   252
      Left            =   5040
      TabIndex        =   285
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   251
      Left            =   4680
      TabIndex        =   284
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   250
      Left            =   4320
      TabIndex        =   283
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   249
      Left            =   3960
      TabIndex        =   282
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   248
      Left            =   3600
      TabIndex        =   281
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   247
      Left            =   3120
      TabIndex        =   280
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   246
      Left            =   2760
      TabIndex        =   279
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   245
      Left            =   2400
      TabIndex        =   278
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   244
      Left            =   2040
      TabIndex        =   277
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   243
      Left            =   1680
      TabIndex        =   276
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   242
      Left            =   1320
      TabIndex        =   275
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   241
      Left            =   960
      TabIndex        =   274
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   240
      Left            =   600
      TabIndex        =   273
      TabStop         =   0   'False
      Top             =   6000
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   239
      Left            =   6120
      TabIndex        =   272
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   238
      Left            =   5760
      TabIndex        =   271
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   237
      Left            =   5400
      TabIndex        =   270
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   236
      Left            =   5040
      TabIndex        =   269
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   235
      Left            =   4680
      TabIndex        =   268
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   234
      Left            =   4320
      TabIndex        =   267
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   233
      Left            =   3960
      TabIndex        =   266
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   232
      Left            =   3600
      TabIndex        =   265
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   231
      Left            =   3120
      TabIndex        =   264
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   230
      Left            =   2760
      TabIndex        =   263
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   229
      Left            =   2400
      TabIndex        =   262
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   228
      Left            =   2040
      TabIndex        =   261
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   227
      Left            =   1680
      TabIndex        =   260
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   226
      Left            =   1320
      TabIndex        =   259
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   225
      Left            =   960
      TabIndex        =   258
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   224
      Left            =   600
      TabIndex        =   257
      TabStop         =   0   'False
      Top             =   5640
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   223
      Left            =   6120
      TabIndex        =   256
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   222
      Left            =   5760
      TabIndex        =   255
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   221
      Left            =   5400
      TabIndex        =   254
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   220
      Left            =   5040
      TabIndex        =   253
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   219
      Left            =   4680
      TabIndex        =   252
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   218
      Left            =   4320
      TabIndex        =   251
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   217
      Left            =   3960
      TabIndex        =   250
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   216
      Left            =   3600
      TabIndex        =   249
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   215
      Left            =   3120
      TabIndex        =   248
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   214
      Left            =   2760
      TabIndex        =   247
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   213
      Left            =   2400
      TabIndex        =   246
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   212
      Left            =   2040
      TabIndex        =   245
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   211
      Left            =   1680
      TabIndex        =   244
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   210
      Left            =   1320
      TabIndex        =   243
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   209
      Left            =   960
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   208
      Left            =   600
      TabIndex        =   241
      TabStop         =   0   'False
      Top             =   5280
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   207
      Left            =   6120
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   206
      Left            =   5760
      TabIndex        =   239
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   205
      Left            =   5400
      TabIndex        =   238
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   204
      Left            =   5040
      TabIndex        =   237
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   203
      Left            =   4680
      TabIndex        =   236
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   202
      Left            =   4320
      TabIndex        =   235
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   201
      Left            =   3960
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   200
      Left            =   3600
      TabIndex        =   233
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   199
      Left            =   3120
      TabIndex        =   232
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   198
      Left            =   2760
      TabIndex        =   231
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   197
      Left            =   2400
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   196
      Left            =   2040
      TabIndex        =   229
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   195
      Left            =   1680
      TabIndex        =   228
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   194
      Left            =   1320
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   193
      Left            =   960
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   192
      Left            =   600
      TabIndex        =   225
      TabStop         =   0   'False
      Top             =   4920
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   191
      Left            =   6120
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   190
      Left            =   5760
      TabIndex        =   223
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   189
      Left            =   5400
      TabIndex        =   222
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   188
      Left            =   5040
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   187
      Left            =   4680
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   186
      Left            =   4320
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   185
      Left            =   3960
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   184
      Left            =   3600
      TabIndex        =   217
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   183
      Left            =   3120
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   182
      Left            =   2760
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   181
      Left            =   2400
      TabIndex        =   214
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   180
      Left            =   2040
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   179
      Left            =   1680
      TabIndex        =   212
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   178
      Left            =   1320
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   177
      Left            =   960
      TabIndex        =   210
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   176
      Left            =   600
      TabIndex        =   209
      TabStop         =   0   'False
      Top             =   4560
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   175
      Left            =   6120
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   174
      Left            =   5760
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   173
      Left            =   5400
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   172
      Left            =   5040
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   171
      Left            =   4680
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   170
      Left            =   4320
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   4200
      Width           =   350
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   169
      Left            =   3960
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   168
      Left            =   3600
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   167
      Left            =   3120
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   166
      Left            =   2760
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   165
      Left            =   2400
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   164
      Left            =   2040
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   163
      Left            =   1680
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   162
      Left            =   1320
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   161
      Left            =   960
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   160
      Left            =   600
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   4200
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   159
      Left            =   6120
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   158
      Left            =   5760
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   157
      Left            =   5400
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   156
      Left            =   5040
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   155
      Left            =   4680
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   154
      Left            =   4320
      TabIndex        =   187
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   153
      Left            =   3960
      TabIndex        =   186
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   152
      Left            =   3600
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   151
      Left            =   3120
      TabIndex        =   184
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   150
      Left            =   2760
      TabIndex        =   183
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   149
      Left            =   2400
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   148
      Left            =   2040
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   147
      Left            =   1680
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   146
      Left            =   1320
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   145
      Left            =   960
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   144
      Left            =   600
      TabIndex        =   177
      TabStop         =   0   'False
      Top             =   3840
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   143
      Left            =   6120
      TabIndex        =   176
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   142
      Left            =   5760
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   141
      Left            =   5400
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   140
      Left            =   5040
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   139
      Left            =   4680
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   138
      Left            =   4320
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   137
      Left            =   3960
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   136
      Left            =   3600
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   135
      Left            =   3120
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   134
      Left            =   2760
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   133
      Left            =   2400
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   132
      Left            =   2040
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   131
      Left            =   1680
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   130
      Left            =   1320
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   129
      Left            =   960
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   128
      Left            =   600
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   3480
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   127
      Left            =   6120
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   126
      Left            =   5760
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   125
      Left            =   5400
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   124
      Left            =   5040
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   123
      Left            =   4680
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   122
      Left            =   4320
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   121
      Left            =   3960
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   120
      Left            =   3600
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   119
      Left            =   3120
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   118
      Left            =   2760
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   117
      Left            =   2400
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   116
      Left            =   2040
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   115
      Left            =   1680
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   114
      Left            =   1320
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   113
      Left            =   960
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   112
      Left            =   600
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   3120
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   111
      Left            =   6120
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   110
      Left            =   5760
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   109
      Left            =   5400
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   108
      Left            =   5040
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   107
      Left            =   4680
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   106
      Left            =   4320
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   105
      Left            =   3960
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   104
      Left            =   3600
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   103
      Left            =   3120
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   102
      Left            =   2760
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   101
      Left            =   2400
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   100
      Left            =   2040
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   99
      Left            =   1680
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   98
      Left            =   1320
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   97
      Left            =   960
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   96
      Left            =   600
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   2760
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   95
      Left            =   6120
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   94
      Left            =   5760
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   93
      Left            =   5400
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   92
      Left            =   5040
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   91
      Left            =   4680
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   90
      Left            =   4320
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   89
      Left            =   3960
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   88
      Left            =   3600
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   87
      Left            =   3120
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   86
      Left            =   2760
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   85
      Left            =   2400
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   84
      Left            =   2040
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   83
      Left            =   1680
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   82
      Left            =   1320
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   81
      Left            =   960
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   80
      Left            =   600
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   2400
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   79
      Left            =   6120
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   78
      Left            =   5760
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   77
      Left            =   5400
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   76
      Left            =   5040
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   75
      Left            =   4680
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   74
      Left            =   4320
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   73
      Left            =   3960
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   72
      Left            =   3600
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   71
      Left            =   3120
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   70
      Left            =   2760
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   69
      Left            =   2400
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   68
      Left            =   2040
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   67
      Left            =   1680
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   66
      Left            =   1320
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   65
      Left            =   960
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   64
      Left            =   600
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   2040
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   63
      Left            =   6120
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   62
      Left            =   5760
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   61
      Left            =   5400
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   60
      Left            =   5040
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   59
      Left            =   4680
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   58
      Left            =   4320
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   57
      Left            =   3960
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   56
      Left            =   3600
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   55
      Left            =   3120
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   54
      Left            =   2760
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   53
      Left            =   2400
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   52
      Left            =   2040
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   51
      Left            =   1680
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   50
      Left            =   1320
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   49
      Left            =   960
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   48
      Left            =   600
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1680
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   32
      Left            =   600
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1320
      Width           =   360
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
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
      Height          =   285
      Index           =   33
      Left            =   960
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1320
      Width           =   360
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   34
      Left            =   1320
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   35
      Left            =   1680
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   36
      Left            =   2040
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   37
      Left            =   2400
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   38
      Left            =   2760
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   39
      Left            =   3120
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   40
      Left            =   3600
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   41
      Left            =   3960
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   42
      Left            =   4320
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   43
      Left            =   4680
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   44
      Left            =   5040
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   45
      Left            =   5400
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   46
      Left            =   5760
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   47
      Left            =   6120
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1320
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   16
      Left            =   600
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   17
      Left            =   960
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   18
      Left            =   1320
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   19
      Left            =   1680
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   21
      Left            =   2400
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   22
      Left            =   2760
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   23
      Left            =   3120
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   24
      Left            =   3600
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   25
      Left            =   3960
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   26
      Left            =   4320
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   27
      Left            =   4680
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   28
      Left            =   5040
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   29
      Left            =   5400
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   30
      Left            =   5760
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   31
      Left            =   6120
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   960
      Width           =   340
   End
   Begin VB.ListBox EEPromList 
      Height          =   255
      Left            =   6120
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   15
      Left            =   6120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   14
      Left            =   5760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   13
      Left            =   5400
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   12
      Left            =   5040
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   11
      Left            =   4680
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   10
      Left            =   4320
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   3600
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   600
      MaxLength       =   2
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   570
      Width           =   340
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "WRITE"
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
      Index           =   5
      Left            =   6600
      TabIndex        =   302
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "LOAD"
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
      Index           =   4
      Left            =   6600
      TabIndex        =   301
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   1680
      TabIndex        =   300
      Top             =   6645
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "SAVE"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   6600
      TabIndex        =   298
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Discrepancies:"
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
      Height          =   195
      Left            =   120
      TabIndex        =   294
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "HEX"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   6600
      TabIndex        =   292
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "EXIT"
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
      Index           =   2
      Left            =   6600
      TabIndex        =   291
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "ASCII"
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
      Index           =   0
      Left            =   6600
      TabIndex        =   290
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   289
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   6165
      TabIndex        =   47
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   5805
      TabIndex        =   46
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   5445
      TabIndex        =   45
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   5085
      TabIndex        =   44
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   4725
      TabIndex        =   43
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "XA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   4365
      TabIndex        =   42
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   4005
      TabIndex        =   41
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   3645
      TabIndex        =   40
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   3165
      TabIndex        =   39
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2805
      TabIndex        =   38
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2445
      TabIndex        =   37
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2085
      TabIndex        =   36
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1725
      TabIndex        =   35
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1365
      TabIndex        =   34
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1005
      TabIndex        =   33
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblIndex 
      Caption         =   "80F0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   16
      Top             =   6040
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "80E0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   15
      Top             =   5680
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "80D0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   14
      Top             =   5320
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "80C0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   13
      Top             =   4960
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "80B0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   12
      Top             =   4600
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "80A0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   4240
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8090"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   3880
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8080"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   3520
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8070"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   3160
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8060"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   2800
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8050"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   2440
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8040"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2080
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8030"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1720
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8020"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1360
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8010"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1000
      Width           =   495
   End
   Begin VB.Label lblIndex 
      Caption         =   "8000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Blue = &H800000
Const Red = &HC0
Const White = &HFFFFFF
Const Gray = &HC0C0C0
Const Black = &H808080
Dim Discp As Integer


Private Sub Command1_Click()
    Init = False
    
    Me.Hide
    Form1.Show
End Sub

Private Sub Form_Activate()
    Dim Index As Integer
    Dim I As Integer
    Dim Temp As String
    Dim EEProm As String
    Dim Bin As String
    Dim J As Integer
    
    If Init <> True Then
        Exit Sub
    ElseIf EEPromList.ListCount = 0 Then
        If UpdateList.ListCount = 0 Then
            EEPromList.Clear
            UpdateList.Clear
            EEProm = ""
            Temp = "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
            For Index = 0 To 255
                Form2.EEPromList.AddItem Temp
                Form2.UpdateList.AddItem Temp
                EEProm = EEProm + Temp + Chr$(13) + Chr$(10)
            Next Index
        End If
    End If
    Form_Initialize
    If EEPromList.ListCount > 0 Then
        For Index = 0 To 15
            EEPromList.ListIndex = Index
            Temp = EEPromList.Text
            For I = 1 To Len(Temp)
                Text(Counter).Text = Mid$(Temp, I, 2)
                I = I + 2
                Counter = Counter + 1
            Next I
        Next Index
        OldLine = 0
        
    End If
    Init = False
    Text(0).BackColor = Black
    Text(0).ForeColor = Blue
    If Text(0).Text = "" Then
        Label(0).Enabled = False
        Label(1).Enabled = False
    End If
    Discp = 0
    
    If EEPromList.ListCount = 0 Then
        Exit Sub
    End If
    If BinList.ListCount = 0 Then
        Exit Sub
    End If
    For I = 0 To 255
        EEPromList.ListIndex = I
        BinList.ListIndex = I
        EEProm = EEPromList.Text
        Bin = BinList.Text
        For J = 1 To Len(EEProm) Step 3
            If Mid$(Bin, J, 1) = "-" Then
                'J = J + 2
            Else
                If UCase(Mid$(Bin, J, 2)) <> UCase(Mid$(EEProm, J, 2)) Then
                    Discp = Discp + 1
                End If
            End If
        Next J
    Next I
    
    Text2.Text = Discp
    Disc
    Text(0).BackColor = Black
    Text(0).ForeColor = Blue
    Label(3).Enabled = True
    
End Sub

Private Sub Form_Initialize()
    Dim I As Integer
    
    For I = 0 To 255
        Text(I).FontBold = True
        Text(I).Alignment = 2
        Text(I).Width = 370
        Text(I).TabStop = False
        Text(I).BackColor = Gray
    Next
    Init = True
    LineNo = 0
    RowNo = 0
    TextBox = 0
    Me.KeyPreview = True
    FirstKey = True
    Me.Caption = MagiString + " © 2000 " + " --- EEProm"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Temp1 As String
    Dim Correct As String
    Dim OL As String
    Dim NL As String
    
    Select Case KeyCode
        Case 39 ' Right Arrow
            RowNo = RowNo + 1
            MoveCursor
        
        Case 37 ' Left Arrow
            RowNo = RowNo - 1
            MoveCursor
                
        Case 34 ' Page Down
            LineNo = LineNo + 16
            MoveCursor
            
        Case 33 ' Page Up
            LineNo = LineNo - 16
            MoveCursor
            
        Case 40 ' Down Arrow
            LineNo = LineNo + 1
            MoveCursor
            
        Case 38 ' Up Arrow
            LineNo = LineNo - 1
            MoveCursor
        
        Case 36 ' Home
            LineNo = 0
            RowNo = 0
            OldLine = 0
            MoveCursor
            Redisplay
            
        Case 35
            RowNo = 15
            OldLine = 240
            LineNo = 255
            MoveCursor
            
        Case 48 To 57 ' numbers 0 - 9
            If FirstKey = True Then
                Text(TextBox).Text = Chr(KeyCode)
                FirstKey = False
                                
            Else
                Text(TextBox).Text = Text(TextBox).Text + Chr(KeyCode)
                FirstKey = True
                EEPromList.ListIndex = OldLine + LineNo
                Temp1 = EEPromList.Text
                Correct = Text(TextBox)
                Mid$(Temp1, (RowNo * 3) + 1, 2) = Correct
                EEPromList.RemoveItem OldLine + LineNo
                EEPromList.AddItem Temp1, OldLine + LineNo
                EEPromList.ListIndex = OldLine + LineNo
                Temp1 = EEPromList.Text
                OL = Mid$(Temp1, (RowNo * 3) + 1, 2)
                BinList.ListIndex = OldLine + LineNo
                Temp1 = BinList.Text
                NL = Mid$(Temp1, (RowNo * 3) + 1, 2)
                If txtCorrect.Visible = True Then
                    If NL = OL Then
                        Discp = Discp - 1
                        Text2.Text = Discp
                        Text(TextBox).ForeColor = Blue
                        txtCorrect.Visible = False
                    End If
                ElseIf NL <> OL Then
                        Discp = Discp + 1
                        Text2.Text = Discp
                        Text(TextBox).ForeColor = Red
                        txtCorrect.Visible = True
                        txtCorrect.Text = NL
                End If
                Discrepancies
                Call Form_KeyUp(39, 0)
            End If
        
        Case 65 To 70 ' alpha A - F
            If FirstKey = True Then
                Text(TextBox).Text = Chr(KeyCode)
                FirstKey = False
            Else
                Text(TextBox).Text = Text(TextBox).Text + Chr(KeyCode)
                FirstKey = True
                EEPromList.ListIndex = OldLine + LineNo
                Temp1 = EEPromList.Text
                Correct = Text(TextBox)
                Mid$(Temp1, (RowNo * 3) + 1, 2) = Correct
                EEPromList.RemoveItem OldLine + LineNo
                EEPromList.AddItem Temp1, OldLine + LineNo
                EEPromList.ListIndex = OldLine + LineNo
                Discrepancies
            End If
        
    End Select
    Text1.Text = ""
End Sub

Private Sub Form_Load()
    Counter = 0
    Me.Hide
End Sub

Private Sub Label_Click(Index As Integer)
    Dim I As Integer
    Dim Temp As String
    
    Select Case Index
        Case 0
            Call ToAscii
            
        Case 2
            Command1_Click
        
        Case 1
            Call ToHex
            
        Case 3
            Call SaveImage
            
        Case 4
            Call LoadImage
        
        Case 5
            Call WriteImage
            
    End Select
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim I As Integer
    
    Label(Index).ForeColor = &H80

    For I = 0 To 5
        If I <> Index Then
            If Label(I).ForeColor = &H80 Then
                Label(I).ForeColor = &H800000
            End If
        End If
    Next I
End Sub

Sub ToAscii()
    Dim I As Integer
    Dim Temp As String
    Dim Hold As Integer

    
    For I = 0 To 255
        Temp = "&H" + Text(I).Text
        Hold = CInt(Temp)
        If Hold = 0 Then
            Text(I).Text = "."
        Else
            Text(I).Text = Chr(Hold)
        End If
    Next I
    Label(0).Enabled = False
    Label(1).Enabled = True
End Sub

Sub ToHex()
    Dim I As Integer
    Dim Temp As String
    Dim Hold As Integer
    
    For I = 0 To 255
        If Text(I).Text = "." Then
            Hold = 0
        Else
            Hold = Asc(Text(I).Text)
        End If
        If Hold < &H10 Then
            Temp = "0" + Hex$(Hold)
        Else
            Temp = Hex$(Hold)
        End If
        Text(I).Text = Temp
    Next I
    Label(0).Enabled = True
    Label(1).Enabled = False
End Sub

Sub MoveCursor()
    Dim Temp  As Integer
    Dim Temp1 As String
    Dim Correct As String
    
    'change previous textbox to default
    Text(TextBox).BackColor = Gray
    If Text(TextBox).ForeColor <> Red Then
        'Text(TextBox).ForeColor = Black
    End If
    'test row going off right
    If RowNo = 16 Then
        LineNo = LineNo + 1
        If LineNo < 16 Then
            RowNo = 0
        Else
            'increment counter redisplay
            RowNo = RowNo - 1
        End If
    End If
    'test row going off left
    If RowNo = -1 Then
        LineNo = LineNo - 1
        If LineNo > 0 Then
            RowNo = 15
        Else
            'decrement counter redisplay
            RowNo = RowNo + 1
        End If
    End If
    'test line going off bottom
    If LineNo > 15 Then
        Temp = LineNo - 15
        'check present position
        LineNo = 15
        'increment counter to redisplay screen
        OldLine = OldLine + Temp
        Redisplay
    End If
    'test line going off top
    If LineNo < 0 Then
        Temp = 0 + LineNo
        LineNo = 0
        'decrement counter to reddisplay screen
        OldLine = OldLine + Temp
        Redisplay
    End If
    
    'rem now check differences
    Disc
    'change textbox color to highlight
    'check first to see if red
    TextBox = (LineNo * 16) + RowNo
    Text(TextBox).BackColor = Black
    If Text(TextBox).ForeColor = Red Then
        txtCorrect.Visible = True
        BinList.ListIndex = OldLine + LineNo
        Temp1 = BinList.Text
        Correct = Mid$(Temp1, (RowNo * 3) + 1, 2)
        txtCorrect.Text = Correct
    Else
        txtCorrect.Visible = False
    End If
End Sub

Sub Redisplay()
    Dim I As Integer
    Dim Index As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim NewTemp As String
    Dim try As Integer
    Dim J As Integer
    
    Temp1 = "800080108020803080408050806080708080809080A080B080C080D080E080F0"
    Temp1 = Temp1 + "810081108120813081408150816081708180819081A081B081C081D081E081F0"
    Temp1 = Temp1 + "820082108220823082408250826082708280829082A082B082C082D082E082F0"
    Temp1 = Temp1 + "830083108320833083408350836083708380839083A083B083C083D083E083F0"
    Temp1 = Temp1 + "840084108420843084408450846084708480849084A084B084C084D084E084F0"
    Temp1 = Temp1 + "850085108520853085408550856085708580859085A085B085C085D085E085F0"
    Temp1 = Temp1 + "860086108620863086408650866086708680869086A086B086C086D086E086F0"
    Temp1 = Temp1 + "870087108720873087408750876087708780879087A087B087C087D087E087F0"
    Temp1 = Temp1 + "880088108820883088408850886088708880889088A088B088C088D088E088F0"
    Temp1 = Temp1 + "890089108920893089408950896089708980899089A089B089C089D089E089F0"
    Temp1 = Temp1 + "8A008A108A208A308A408A508A608A708A808A908AA08AB08AC08AD08AE08AF0"
    Temp1 = Temp1 + "8B008B108B208B308B408B508B608B708B808B908BA08BB08BC08BD08BE08BF0"
    Temp1 = Temp1 + "8C008C108C208C308C408C508C608C708C808C908CA08CB08CC08CD08CE08CF0"
    Temp1 = Temp1 + "8D008D108D208D308D408D508D608D708D808D908DA08DB08DC08DD08DE08DF0"
    Temp1 = Temp1 + "8E008E108E208E308E408E508E608E708E808E908EA08EB08EC08ED08EE08EF0"
    Temp1 = Temp1 + "8F008F108F208F308F408F508F608F708F808F908FA08FB08FC08FD08FE08FF0"
    
    Counter = 0
    J = 0
    If OldLine < 0 Then
        OldLine = 0
    End If
    If OldLine > 240 Then
        OldLine = 240
    End If
    
    If EEPromList.ListCount > 0 Then
        For Index = OldLine To OldLine + 15
            EEPromList.ListIndex = Index
            Temp = EEPromList.Text
            For I = 1 To Len(Temp)
                Text(Counter).Text = Mid$(Temp, I, 2)
                I = I + 2
                Counter = Counter + 1
            Next I
        try = (Index * 4) + 1
        NewTemp = Mid$(Temp1, try, 4)
        lblIndex(J).Caption = NewTemp
        J = J + 1
        Next Index
    End If
    If Label(1).Enabled = True Then
        ToAscii
    End If
    
End Sub

Sub Disc()
    Dim I As Integer
    Dim J As Integer
    Dim cnt As Integer
    Dim Bin As String
    Dim Temp As String
    Dim Temp1 As String
    Dim ConInt As Integer
    Dim ConTemp1 As Integer
    
    cnt = 0
    For I = OldLine To OldLine + 15
        BinList.ListIndex = I
        Bin = BinList.Text
        For J = 1 To Len(Bin) Step 3
            If Mid$(Bin, J, 1) = "-" Then
                Text(cnt).ForeColor = Blue
            ElseIf Label(0).Enabled = True Then
                If UCase(Mid$(Bin, J, 2)) <> Text(cnt).Text Then
                    Text(cnt).ForeColor = Red
                Else
                    Text(cnt).ForeColor = Blue
                End If
            ElseIf Label(0).Enabled = False Then
                'Temp = "&H" + Mid$(Bin, J, 2)
                Temp = Mid$(Bin, J, 2)
                'ConInt = CInt(Temp)
                Temp1 = Text(cnt).Text
                If Temp1 = "." Then
                    ConTemp1 = 0
                Else
                    'ConTemp1 = Asc(Temp1)
                End If
                'If ConInt <> ConTemp1 Then
                If (Temp) <> (Temp1) Then
                    Text(cnt).ForeColor = Red
                Else
                    Text(cnt).ForeColor = Blue
                End If
            
            End If
        cnt = cnt + 1
        Next J
    Next I
End Sub

'*********************************
'*
'remember to add to for e3 or 2d
'*
Sub SaveImage()
    Dim Hdl As Integer
    Dim I As Integer
    Dim J As Integer
    Dim cnt As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim Test As Integer
    Dim Hold As String
    Dim Test1 As String
    
    Hdl = FreeFile

    If EEPromList.ListCount = 0 Then
        Exit Sub
    End If
    On Error GoTo Handler
    cnt = 1
    
    Temp1 = "C:\Magic\BINs\"
    Temp = InputBox("Save EEprom Image as FileName", MagiString + " © 2000 -- Save EEProm Image")
    I = InStr(1, Temp, ".")
    If I = 0 Then
        Temp = Temp + ".BIN"
    Else
        Temp = Left$(Temp, I - 1) + ".BIN"
    End If
    Temp1 = Temp1 + Temp
    
    Open Temp1 For Binary Access Write As #Hdl
    
    For I = 0 To 255
        EEPromList.ListIndex = I
        Temp = EEPromList.Text
        For J = 1 To Len(Temp) Step 3
            Temp1 = Mid$(Temp, J, 2)
            Hold = Left$(Temp1, 1)
            Select Case Hold
                Case "0"
                    Test = 0
                Case "1"
                    Test = 16
                Case "2"
                    Test = 32
                Case "3"
                    Test = 48
                Case "4"
                    Test = 64
                Case "5"
                    Test = 80
                Case "6"
                    Test = 96
                Case "7"
                    Test = 112
                Case "8"
                    Test = 128
                Case "9"
                    Test = 144
                Case "A"
                    Test = 160
                Case "B"
                    Test = 176
                Case "C"
                    Test = 192
                Case "D"
                    Test = 208
                Case "E"
                    Test = 224
                Case "F"
                    Test = 240
            End Select
            Hold = Right$(Temp1, 1)
            Select Case Hold
                Case "1"
                    Test = Test + 1
                Case "2"
                    Test = Test + 2
                Case "3"
                    Test = Test + 3
                Case "4"
                    Test = Test + 4
                Case "5"
                    Test = Test + 5
                Case "6"
                    Test = Test + 6
                Case "7"
                    Test = Test + 7
                Case "8"
                    Test = Test + 8
                Case "9"
                    Test = Test + 9
                Case "A"
                    Test = Test + 10
                Case "B"
                    Test = Test + 11
                Case "C"
                    Test = Test + 12
                Case "D"
                    Test = Test + 13
                Case "E"
                    Test = Test + 14
                Case "F"
                    Test = Test + 15
            End Select
            Test1 = Test1 + Chr(Test)
            
            If cnt > 4096 Then
                Exit For
            End If
        Next J
        Put #Hdl, cnt, Test1
        cnt = cnt + 16
        Test1 = ""
    Next I
    txtStat.Text = "File save successfully."
    Close #Hdl
    Exit Sub
    
Handler:

    Close #Hdl
    txtStat.Text = "Failed saving file."
    
End Sub

Sub LoadImage()
    Dim Hdl As Integer
    Dim Msg As String
    Dim cnt As Integer
    Dim Temp As String * 16
    Dim I As Integer
    Dim J As Integer
    Dim Hold As Integer
    Dim Test1 As String
    
    BinFile = ""
    FilName = "C:\Magic\Bins\"
    Form3.FileList.Pattern = "*.BIN"
    Form3.Show vbModal
    If BinFile = "" Then
        Exit Sub
    End If
    
    Hdl = FreeFile
    Msg = "Unable to open file."
    cnt = 1
    
    'On Error GoTo Handler
    Open "C:\Magic\Bins\" + BinFile For Binary Access Read As #Hdl Len = 16

        Get #Hdl, cnt, Temp
        Msg = ""
        Hold = Asc(Mid$(Temp, 1, 1))
        If Hold <> 51 Then
            Msg = "Invalid Bin file."
            GoTo Handler
        End If
        
        Form2.EEPromList.Clear
        'Form2.UpdateList.Clear
        
        For I = 1 To 256
            Test1 = ""
            Get #1, cnt, Temp
            For J = 1 To 16
                Hold = Asc(Mid$(Temp, J, 1))
                If Hold < &H10 Then
                    Test1 = Test1 + "0" + Hex(Hold) + " "
                Else
                    Test1 = Test1 + Hex(Hold) + " "
                End If
            Next J
            cnt = cnt + 16
            Form2.EEPromList.AddItem Test1
            'Form2.UpdateList.AddItem Test1
        Next I
        txtStat.Text = "Image loaded successfully."
        Close #Hdl
        Counter = 0
        Init = True
        Form_Activate
        Msg = "Bin image loaded successfully."

Handler:

    txtStat.Text = Msg
    Close #Hdl
    
End Sub

Sub Discrepancies()
    Dim I As Integer
    Dim J As Integer
    Dim EEProm As String
    Dim Bin As String
    
    Discp = 0
    
    For I = 0 To 255
        EEPromList.ListIndex = I
        BinList.ListIndex = I
        EEProm = EEPromList.Text
        Bin = BinList.Text
        For J = 1 To Len(EEProm) Step 3
            If Mid$(Bin, J, 1) = "-" Then
                'J = J + 2
            Else
                If UCase(Mid$(Bin, J, 2)) <> UCase(Mid$(EEProm, J, 2)) Then
                    Discp = Discp + 1
                End If
            End If
        Next J
    Next I
    
    Text2.Text = Discp
    Disc
End Sub
