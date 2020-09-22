VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   " Wordpuzzle"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.Slider sldDelay 
      Height          =   375
      Left            =   120
      TabIndex        =   313
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Max             =   5000
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   10920
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Reset"
      Height          =   495
      Left            =   9720
      TabIndex        =   312
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemoveListiem 
      Caption         =   "Remove"
      Height          =   255
      Left            =   9720
      TabIndex        =   311
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add"
      Height          =   255
      Left            =   9720
      TabIndex        =   310
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8040
      TabIndex        =   309
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save to File"
      Height          =   255
      Left            =   9720
      TabIndex        =   307
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load from File"
      Height          =   255
      Left            =   9720
      TabIndex        =   306
      Top             =   720
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   1815
      Left            =   5040
      TabIndex        =   305
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save from table"
      Height          =   255
      Left            =   9720
      TabIndex        =   303
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load into table"
      Height          =   255
      Left            =   9720
      TabIndex        =   302
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   299
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   301
      Text            =   "E"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   298
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   300
      Text            =   "S"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   297
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   299
      Text            =   "N"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   296
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   298
      Text            =   "E"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   295
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   297
      Text            =   "D"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   294
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   296
      Text            =   "O"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   293
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   295
      Text            =   "B"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   292
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   294
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   291
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   293
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   290
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   292
      Text            =   "E"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   289
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   291
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   288
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   290
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   287
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   289
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   286
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   288
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   287
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   284
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   286
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   283
      Left            =   840
      MaxLength       =   1
      TabIndex        =   285
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   282
      Left            =   600
      MaxLength       =   1
      TabIndex        =   284
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   281
      Left            =   360
      MaxLength       =   1
      TabIndex        =   283
      Text            =   "M"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   280
      Left            =   120
      MaxLength       =   1
      TabIndex        =   282
      Text            =   "B"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   279
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   281
      Text            =   "B"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   278
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   280
      Text            =   "O"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   277
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   279
      Text            =   "D"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   276
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   278
      Text            =   "E"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   275
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   277
      Text            =   "N"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   274
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   276
      Text            =   "S"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   273
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   275
      Text            =   "E"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   272
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   274
      Text            =   "E"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   271
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   273
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   270
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   272
      Text            =   "E"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   269
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   271
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   268
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   270
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   267
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   269
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   266
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   268
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   265
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   267
      Text            =   "E"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   264
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   266
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   263
      Left            =   840
      MaxLength       =   1
      TabIndex        =   265
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   262
      Left            =   600
      MaxLength       =   1
      TabIndex        =   264
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   261
      Left            =   360
      MaxLength       =   1
      TabIndex        =   263
      Text            =   "O"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   260
      Left            =   120
      MaxLength       =   1
      TabIndex        =   262
      Text            =   "M"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   259
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   261
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   258
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   260
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   257
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   259
      Text            =   "B"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   256
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   258
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   255
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   257
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   254
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   256
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   253
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   255
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   252
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   254
      Text            =   "E"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   251
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   253
      Text            =   "E"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   250
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   252
      Text            =   "S"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   249
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   251
      Text            =   "N"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   248
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   250
      Text            =   "E"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   247
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   249
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   246
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   248
      Text            =   "E"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   245
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   247
      Text            =   "B"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   244
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   246
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   243
      Left            =   840
      MaxLength       =   1
      TabIndex        =   245
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   242
      Left            =   600
      MaxLength       =   1
      TabIndex        =   244
      Text            =   "D"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   241
      Left            =   360
      MaxLength       =   1
      TabIndex        =   243
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   240
      Left            =   120
      MaxLength       =   1
      TabIndex        =   242
      Text            =   "M"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   239
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   241
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   238
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   240
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   237
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   239
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   236
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   238
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   235
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   237
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   234
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   236
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   233
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   235
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   232
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   234
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   231
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   233
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   230
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   232
      Text            =   "N"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   229
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   231
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   228
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   230
      Text            =   "E"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   227
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   229
      Text            =   "S"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   226
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   228
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   225
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   227
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   224
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   226
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   223
      Left            =   840
      MaxLength       =   1
      TabIndex        =   225
      Text            =   "E"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   222
      Left            =   600
      MaxLength       =   1
      TabIndex        =   224
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   221
      Left            =   360
      MaxLength       =   1
      TabIndex        =   223
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   220
      Left            =   120
      MaxLength       =   1
      TabIndex        =   222
      Text            =   "M"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   219
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   221
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   218
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   220
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   217
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   219
      Text            =   "E"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   216
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   218
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   215
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   217
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   214
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   216
      Text            =   "E"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   213
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   215
      Text            =   "E"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   212
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   214
      Text            =   "S"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   211
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   213
      Text            =   "N"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   210
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   212
      Text            =   "E"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   209
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   211
      Text            =   "D"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   208
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   210
      Text            =   "N"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   207
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   209
      Text            =   "E"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   206
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   208
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   205
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   207
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   204
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   206
      Text            =   "N"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   203
      Left            =   840
      MaxLength       =   1
      TabIndex        =   205
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   202
      Left            =   600
      MaxLength       =   1
      TabIndex        =   204
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   201
      Left            =   360
      MaxLength       =   1
      TabIndex        =   203
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   200
      Left            =   120
      MaxLength       =   1
      TabIndex        =   202
      Text            =   "M"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   199
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   201
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   198
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   200
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   197
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   199
      Text            =   "S"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   196
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   198
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   195
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   197
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   194
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   196
      Text            =   "D"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   193
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   195
      Text            =   "O"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   192
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   194
      Text            =   "B"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   191
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   193
      Text            =   "M"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   190
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   192
      Text            =   "D"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   189
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   191
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   188
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   190
      Text            =   "M"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   187
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   189
      Text            =   "B"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   186
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   188
      Text            =   "S"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   185
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   187
      Text            =   "S"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   184
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   186
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   183
      Left            =   840
      MaxLength       =   1
      TabIndex        =   185
      Text            =   "N"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   182
      Left            =   600
      MaxLength       =   1
      TabIndex        =   184
      Text            =   "S"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   181
      Left            =   360
      MaxLength       =   1
      TabIndex        =   183
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   180
      Left            =   120
      MaxLength       =   1
      TabIndex        =   182
      Text            =   "E"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   179
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   181
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   178
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   180
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   177
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   179
      Text            =   "E"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   176
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   178
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   175
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   177
      Text            =   "S"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   174
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   176
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   173
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   175
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   172
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   174
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   171
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   173
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   170
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   172
      Text            =   "D"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   169
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   171
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   168
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   170
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   167
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   169
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   166
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   168
      Text            =   "E"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   165
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   167
      Text            =   "N"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   164
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   166
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   163
      Left            =   840
      MaxLength       =   1
      TabIndex        =   165
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   162
      Left            =   600
      MaxLength       =   1
      TabIndex        =   164
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   161
      Left            =   360
      MaxLength       =   1
      TabIndex        =   163
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   160
      Left            =   120
      MaxLength       =   1
      TabIndex        =   162
      Text            =   "M"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   159
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   161
      Text            =   "B"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   158
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   160
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   157
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   159
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   156
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   158
      Text            =   "E"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   155
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   157
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   154
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   156
      Text            =   "N"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   153
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   155
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   152
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   154
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   151
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   153
      Text            =   "O"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   150
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   152
      Text            =   "B"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   149
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   151
      Text            =   "B"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   148
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   150
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   147
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   149
      Text            =   "E"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   146
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   148
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   145
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   147
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   144
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   146
      Text            =   "E"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   143
      Left            =   840
      MaxLength       =   1
      TabIndex        =   145
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   142
      Left            =   600
      MaxLength       =   1
      TabIndex        =   144
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   141
      Left            =   360
      MaxLength       =   1
      TabIndex        =   143
      Text            =   "M"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   140
      Left            =   120
      MaxLength       =   1
      TabIndex        =   142
      Text            =   "E"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   139
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   141
      Text            =   "O"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   138
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   140
      Text            =   "O"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   137
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   139
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   136
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   138
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   135
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   137
      Text            =   "S"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   134
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   136
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   133
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   135
      Text            =   "E"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   132
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   134
      Text            =   "B"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   131
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   133
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   130
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   132
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   129
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   131
      Text            =   "O"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   128
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   130
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   127
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   129
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   126
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   128
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   125
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   127
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   124
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   126
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   123
      Left            =   840
      MaxLength       =   1
      TabIndex        =   125
      Text            =   "D"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   122
      Left            =   600
      MaxLength       =   1
      TabIndex        =   124
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   121
      Left            =   360
      MaxLength       =   1
      TabIndex        =   123
      Text            =   "M"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   120
      Left            =   120
      MaxLength       =   1
      TabIndex        =   122
      Text            =   "E"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   119
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   121
      Text            =   "D"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   118
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   120
      Text            =   "B"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   117
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   119
      Text            =   "D"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   116
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   118
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   115
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   117
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   114
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   116
      Text            =   "N"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   113
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   115
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   112
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   114
      Text            =   "D"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   111
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   113
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   110
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   112
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   109
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   111
      Text            =   "D"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   108
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   110
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   107
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   109
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   106
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   108
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   105
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   107
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   104
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   106
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   103
      Left            =   840
      MaxLength       =   1
      TabIndex        =   105
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   102
      Left            =   600
      MaxLength       =   1
      TabIndex        =   104
      Text            =   "O"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   101
      Left            =   360
      MaxLength       =   1
      TabIndex        =   103
      Text            =   "M"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   100
      Left            =   120
      MaxLength       =   1
      TabIndex        =   102
      Text            =   "S"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   99
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   101
      Text            =   "E"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   98
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   100
      Text            =   "O"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   97
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   99
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   96
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   98
      Text            =   "E"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   95
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   97
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   94
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   96
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   93
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   95
      Text            =   "E"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   92
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   94
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   91
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   93
      Text            =   "O"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   90
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   92
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   89
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   91
      Text            =   "E"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   88
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   90
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   87
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   89
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   86
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   88
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   85
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   87
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   84
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   86
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   83
      Left            =   840
      MaxLength       =   1
      TabIndex        =   85
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   82
      Left            =   600
      MaxLength       =   1
      TabIndex        =   84
      Text            =   "M"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   81
      Left            =   360
      MaxLength       =   1
      TabIndex        =   83
      Text            =   "B"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   80
      Left            =   120
      MaxLength       =   1
      TabIndex        =   82
      Text            =   "N"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   79
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   81
      Text            =   "N"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   78
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   80
      Text            =   "O"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   77
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   79
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   76
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   78
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   75
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   77
      Text            =   "N"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   74
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   76
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   73
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   75
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   72
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   74
      Text            =   "D"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   71
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   73
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   70
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   72
      Text            =   "B"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   69
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   71
      Text            =   "N"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   68
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   70
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   67
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   69
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   66
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   68
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   65
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   67
      Text            =   "M"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   64
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   66
      Text            =   "E"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   63
      Left            =   840
      MaxLength       =   1
      TabIndex        =   65
      Text            =   "E"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   62
      Left            =   600
      MaxLength       =   1
      TabIndex        =   64
      Text            =   "S"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   61
      Left            =   360
      MaxLength       =   1
      TabIndex        =   63
      Text            =   "N"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   60
      Left            =   120
      MaxLength       =   1
      TabIndex        =   62
      Text            =   "E"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   59
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   61
      Text            =   "S"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   58
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   60
      Text            =   "D"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   57
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   59
      Text            =   "B"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   56
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   58
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   55
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   57
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   54
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   56
      Text            =   "S"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   53
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   55
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   52
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   54
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   51
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   53
      Text            =   "O"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   50
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   52
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   49
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   51
      Text            =   "S"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   48
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   50
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   47
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   49
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   46
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   48
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   45
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   47
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   44
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   46
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   43
      Left            =   840
      MaxLength       =   1
      TabIndex        =   45
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   42
      Left            =   600
      MaxLength       =   1
      TabIndex        =   44
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   41
      Left            =   360
      MaxLength       =   1
      TabIndex        =   43
      Text            =   "M"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   40
      Left            =   120
      MaxLength       =   1
      TabIndex        =   42
      Text            =   "D"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   39
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   41
      Text            =   "E"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   38
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   40
      Text            =   "E"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   37
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   39
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   36
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   38
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   35
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   37
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   34
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   36
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   33
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   35
      Text            =   "E"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   32
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   34
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   31
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   33
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   30
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   32
      Text            =   "B"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   29
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   31
      Text            =   "E"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   28
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   30
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   27
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   29
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   26
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   28
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   25
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   27
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   24
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   23
      Left            =   840
      MaxLength       =   1
      TabIndex        =   25
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   22
      Left            =   600
      MaxLength       =   1
      TabIndex        =   24
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   21
      Left            =   360
      MaxLength       =   1
      TabIndex        =   23
      Text            =   "M"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   20
      Left            =   120
      MaxLength       =   1
      TabIndex        =   22
      Text            =   "O"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find!"
      Height          =   735
      Left            =   2640
      TabIndex        =   21
      Top             =   4560
      Width           =   7815
   End
   Begin VB.ListBox lstStrings 
      Height          =   1230
      ItemData        =   "Form1.frx":01AB
      Left            =   5160
      List            =   "Form1.frx":0248
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   19
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   19
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   18
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   18
      Text            =   "N"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   17
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "B"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   16
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "O"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   15
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "D"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   14
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   13
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "N"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   12
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   11
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   10
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   10
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   9
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   8
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   7
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "F"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   6
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "R"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   5
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   4
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "B"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "O"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "D"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   1
      Left            =   360
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "B"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblword 
      Height          =   255
      Left            =   5160
      TabIndex        =   315
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Delay:"
      Height          =   255
      Left            =   240
      TabIndex        =   314
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Search for:"
      Height          =   255
      Left            =   5160
      TabIndex        =   308
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "The keyword is:"
      Height          =   255
      Left            =   3960
      TabIndex        =   304
      Top             =   3960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   WordPuzzle Solver
'       (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net

Private Sub cmdRemoveListiem_Click()
    For loop1 = 0 To lstStrings.ListCount - 1
        If lstStrings.Selected(loop1) = True Then lstStrings.RemoveItem (loop1)
    Next loop1
End Sub

Private Sub cmdfind_Click()
Dim bytBYTE(100), strString As String
Dim length, i, j As Long
    For xy = 0 To lstStrings.ListCount
        SearchIT lstStrings.List(xy)
    Next xy
    lblword.Caption = ""
For xy = 0 To 299
    If Text1(xy).BackColor <> vbRed Then lblword.Caption = lblword.Caption & Text1(xy).Text
Next xy


End Sub

Private Sub SearchIT(Str As String)
                On Error Resume Next
'MsgBox Len(Str)
Dim FounD As Boolean
For loop1 = 0 To 299
FounD = True
    If Text1(loop1).Text = Mid(Str, 1, 1) Then
    
        If ((GetRow(Int(loop1)) - 1) * 20 + 19 >= loop1 + Len(Str) - 1) Then 'Horizontal right
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop2 + loop1).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then Highlight Int(loop1), Len(Str) - 1
        End If
        FounD = True
        
        If (GetRow(Int(loop1)) - 1) * 10 <= loop1 - Len(Str) Then
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 - loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then Highlight loop1 - Len(Str) + 1, Len(Str) - 1
        End If
        FounD = True
        
        If GetRow(Int(loop1)) + Len(Str) - 1 <= 15 Then
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 + 20 * loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then HighlightUD Int(loop1), Len(Str) - 1
        End If
        FounD = True
        
        If GetRow(Int(loop1)) - Len(Str) + 1 >= 0 Then
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 - 20 * loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then
                For loop3 = 0 To Len(Str) - 1
                    Text1(loop1 - 20 * loop3).BackColor = vbRed
                Next loop3
            End If
        End If
        FounD = True

        'If (GetRow(Int(loop1)) - Len(Str) + 1 >= 0) And ((GetRow(Int(loop1)) - 1) * 20 + 19 >= loop1 + Len(Str) - 1) Then
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 - 20 * loop2 - loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then
                For loop3 = 0 To Len(Str) - 1
                    Text1(loop1 - 20 * loop3 - loop3).BackColor = vbRed
                Next loop3
            End If
        'End If
        FounD = True
        
        'If (GetRow(Int(loop1)) - Len(Str) + 1 >= 0) And ((GetRow(Int(loop1)) - 1) * 20 + 19 >= loop1 + Len(Str) - 1) Then
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 + 20 * loop2 + loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then
                For loop3 = 0 To Len(Str) - 1
                    Text1(loop1 + 20 * loop3 + loop3).BackColor = vbRed
                Next loop3
            End If
        'End If
        FounD = True
        
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 - 20 * loop2 + loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then
                For loop3 = 0 To Len(Str) - 1
                    Text1(loop1 - 20 * loop3 + loop3).BackColor = vbRed
                Next loop3
            End If
        'End If
        FounD = True
        
            For loop2 = 0 To Len(Str) - 1
                If Text1(loop1 + 20 * loop2 - loop2).Text <> Mid(Str, loop2 + 1, 1) Then FounD = False
            Next loop2
            If FounD = True Then
                For loop3 = 0 To Len(Str) - 1
                    Text1(loop1 + 20 * loop3 - loop3).BackColor = vbRed
                Next loop3
            End If
        'End If
        FounD = True
               
        
    End If
    For wait = 1 To sldDelay.Value
        DoEvents
        DoEvents
    Next wait
Next loop1
End Sub

Private Sub HighlightUD(Current As Integer, lenght As Integer)
For loop1 = 0 To lenght
    Text1(Current + loop1 * 20).BackColor = vbRed
Next loop1
End Sub

Private Sub Highlight(Current As Integer, lenght As Integer)
For loop1 = 0 To lenght
    Text1(Current + loop1).BackColor = vbRed
Next loop1
End Sub


Function GetRow(CurrentBox As Integer)
Dim row As Integer
Select Case CurrentBox
    Case 0 To 19:
        row = 1
    Case 20 To 39:
        row = 2
    Case 40 To 59:
        row = 3
    Case 60 To 79:
        row = 4
    Case 80 To 99:
        row = 5
    Case 100 To 119:
        row = 6
    Case 120 To 139:
        row = 7
    Case 140 To 159:
        row = 8
    Case 160 To 179:
        row = 9
    Case 180 To 199:
        row = 10
    Case 200 To 219:
        row = 11
    Case 220 To 239:
        row = 12
    Case 240 To 259:
        row = 13
    Case 260 To 279:
        row = 14
    Case 280 To 299:
        row = 15
End Select
GetRow = row
End Function

Private Sub Command2_Click()
For loop1 = 0 To 299
    Text1(loop1).Text = Mid(rtbText.Text, loop1 + 1, 1)
Next loop1
End Sub

Private Sub Command3_Click()
rtbText.Text = ""
For loop1a = 0 To 299
    rtbText.Text = rtbText.Text & Text1(loop1a).Text
Next loop1a
End Sub

Private Sub Command4_Click()
    On Error GoTo FinaliseError
    
    dlg.CancelError = True
    dlg.Filter = "All RTF Files|*.rtf|"
    dlg.Flags = cdlOFNFileMustExist
    dlg.ShowOpen
    
    If dlg.FileName = "" Then Exit Sub
    
    rtbText.LoadFile dlg.FileName

FinaliseError:
End Sub

Private Sub Command5_Click()
    On Error GoTo FinaliseError
    
    dlg.CancelError = True
    dlg.Filter = "All RTF Files|*.rtf|"
    dlg.Flags = cdlOFNFileMustExist
    dlg.ShowSave
    
    If dlg.FileName = "" Then Exit Sub
    
    
    rtbText.SaveFile dlg.FileName

FinaliseError:
End Sub

Private Sub Command6_Click()
Dim temp As String
For loop1 = 1 To Len(Text2.Text)
    temp = temp & UCase(Mid(Text2.Text, loop1, 1))
Next loop1
    lstStrings.AddItem temp
    Text2.Text = ""
End Sub



Private Sub Command7_Click()
rtbText.Text = ""
lblword.Caption = ""
For loop1 = 0 To 299
    Text1(loop1).Text = ""
    Text1(loop1).BackColor = vbWhite
Next loop1
lstStrings.Clear
End Sub

Private Sub Form_Load()

Command2_Click
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Text1(Index + 1).SetFocus
    Text1(Index + 1).SelStart = 0
    Text1(Index + 1).SelLength = 1
    Text1(Index).Text = UCase(Chr(KeyCode))
End Sub
