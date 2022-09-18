VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTest 
   Caption         =   "ShapeEx control test"
   ClientHeight    =   9432
   ClientLeft      =   228
   ClientTop       =   696
   ClientWidth     =   16704
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9432
   ScaleWidth      =   16704
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboClickable 
      Height          =   336
      ItemData        =   "frmTest.frx":0000
      Left            =   9100
      List            =   "frmTest.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   88
      Top             =   4560
      Width           =   2904
   End
   Begin VB.ComboBox cboQuality 
      Height          =   336
      ItemData        =   "frmTest.frx":0004
      Left            =   9100
      List            =   "frmTest.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   100
      Top             =   6240
      Width           =   2904
   End
   Begin VB.ComboBox cboFlipped 
      Height          =   336
      ItemData        =   "frmTest.frx":0008
      Left            =   9100
      List            =   "frmTest.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   95
      Top             =   5400
      Width           =   2904
   End
   Begin VB.ComboBox cboFillStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":000C
      Left            =   9100
      List            =   "frmTest.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   84
      Top             =   3480
      Width           =   2904
   End
   Begin VB.CommandButton cmdChangeFillColor 
      Caption         =   "Change"
      Height          =   372
      Left            =   9940
      TabIndex        =   82
      Top             =   3036
      Width           =   2052
   End
   Begin VB.CommandButton cmdChangeBorderColor 
      Caption         =   "Change"
      Height          =   372
      Left            =   9940
      TabIndex        =   71
      Top             =   1356
      Width           =   2052
   End
   Begin VB.ComboBox cboBackStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":0010
      Left            =   9100
      List            =   "frmTest.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   960
      Width           =   2904
   End
   Begin VB.ComboBox cboUseSubclassing 
      Height          =   336
      ItemData        =   "frmTest.frx":0014
      Left            =   9100
      List            =   "frmTest.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   113
      Top             =   8336
      Width           =   2904
   End
   Begin VB.CommandButton cmdChangeBackColor 
      Caption         =   "Change"
      Height          =   372
      Left            =   9940
      TabIndex        =   66
      Top             =   516
      Width           =   2052
   End
   Begin VB.ComboBox cboStyle3DEffect 
      Height          =   336
      ItemData        =   "frmTest.frx":0018
      Left            =   9100
      List            =   "frmTest.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   111
      Top             =   7916
      Width           =   2904
   End
   Begin VB.ComboBox cboStyle3D 
      Height          =   336
      ItemData        =   "frmTest.frx":001C
      Left            =   9100
      List            =   "frmTest.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   109
      Top             =   7508
      Width           =   2904
   End
   Begin VB.CommandButton cmdDecreaseBordderWidth 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   11680
      TabIndex        =   77
      Top             =   2400
      Width           =   312
   End
   Begin VB.CommandButton cmdIncreaseBordderWidth 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   11680
      TabIndex        =   76
      Top             =   2220
      Width           =   312
   End
   Begin VB.TextBox txtBorderWidth 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9100
      TabIndex        =   75
      Text            =   "1"
      Top             =   2244
      Width           =   2484
   End
   Begin VB.ComboBox cboBorderStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":0020
      Left            =   9100
      List            =   "frmTest.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   73
      Top             =   1800
      Width           =   2904
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   14512
      Top             =   6720
   End
   Begin VB.CheckBox chkAutoRotation 
      Caption         =   "Animate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   12532
      TabIndex        =   104
      Top             =   6752
      Width           =   968
   End
   Begin VB.Timer tmrAutoRotation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14116
      Top             =   6720
   End
   Begin VB.TextBox txtVertices 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9100
      TabIndex        =   115
      Text            =   "5"
      Top             =   8780
      Width           =   624
   End
   Begin VB.TextBox txtShift 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9100
      TabIndex        =   106
      Text            =   "20"
      Top             =   7096
      Width           =   624
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Left            =   9100
      TabIndex        =   97
      Top             =   5880
      Width           =   2904
      _ExtentX        =   5122
      _ExtentY        =   529
      _Version        =   327682
      Max             =   100
      SelStart        =   100
      Value           =   100
   End
   Begin ComctlLib.Slider sldRotationDegrees 
      Height          =   300
      Left            =   9100
      TabIndex        =   103
      Top             =   6716
      Width           =   2904
      _ExtentX        =   5122
      _ExtentY        =   529
      _Version        =   327682
      Max             =   360
   End
   Begin ComctlLib.Slider sldCurvingFactor 
      Height          =   300
      Left            =   9100
      TabIndex        =   91
      Top             =   4980
      Width           =   2904
      _ExtentX        =   5122
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clickable:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   7404
      TabIndex        =   87
      Top             =   4620
      Width           =   756
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Quality:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   18
      Left            =   7404
      TabIndex        =   99
      Top             =   6300
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flipped:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   17
      Left            =   7404
      TabIndex        =   94
      Top             =   5460
      Width           =   636
   End
   Begin VB.Label Label1 
      Caption         =   "New properties:"
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   13
      Left            =   9100
      TabIndex        =   86
      Top             =   4200
      Width           =   2208
   End
   Begin VB.Label Label5 
      Caption         =   "Only solid y transparent."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   12160
      TabIndex        =   85
      Top             =   3520
      Width           =   2352
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "FillStyle:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   16
      Left            =   7404
      TabIndex        =   83
      Top             =   3540
      Width           =   648
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "FillColor:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   15
      Left            =   7404
      TabIndex        =   80
      Top             =   3120
      Width           =   696
   End
   Begin Proyect1.ShapeEx shpFillColor 
      Height          =   240
      Left            =   9130
      TabIndex        =   81
      Top             =   3120
      Width           =   552
      _ExtentX        =   974
      _ExtentY        =   423
      BackStyle       =   1
      BorderColor     =   8421504
      CurvingFactor   =   30
   End
   Begin VB.Label Label4 
      Caption         =   "Not available in ShapeEx because it is not supported by GDI+."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   9100
      TabIndex        =   79
      Top             =   2710
      Width           =   5112
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DrawMode:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   14
      Left            =   7404
      TabIndex        =   78
      Top             =   2700
      Width           =   924
   End
   Begin Proyect1.ShapeEx shpBorderColor 
      Height          =   240
      Left            =   9130
      TabIndex        =   70
      Top             =   1440
      Width           =   552
      _ExtentX        =   974
      _ExtentY        =   423
      BackStyle       =   1
      BorderColor     =   8421504
      CurvingFactor   =   30
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BorderColor:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   13
      Left            =   7404
      TabIndex        =   69
      Top             =   1440
      Width           =   1008
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BackStyle:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   12
      Left            =   7404
      TabIndex        =   67
      Top             =   1020
      Width           =   792
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "UseSubclassing:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   7
      Left            =   7428
      TabIndex        =   112
      Top             =   8400
      Width           =   1248
   End
   Begin Proyect1.ShapeEx shpBackColor 
      Height          =   240
      Left            =   9130
      TabIndex        =   65
      Top             =   600
      Width           =   550
      _ExtentX        =   974
      _ExtentY        =   423
      BackStyle       =   1
      BorderColor     =   8421504
      CurvingFactor   =   30
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BackColor:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   7404
      TabIndex        =   64
      Top             =   600
      Width           =   840
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   5
      Left            =   5940
      Shape           =   5  'Rounded Square
      Top             =   600
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   4
      Left            =   4860
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   3
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   600
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   2
      Left            =   2700
      Shape           =   2  'Oval
      Top             =   600
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   1
      Left            =   1620
      Shape           =   1  'Square
      Top             =   600
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   552
      Index           =   0
      Left            =   540
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Original Shape control:"
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   12
      Left            =   504
      TabIndex        =   0
      Top             =   180
      Width           =   2208
   End
   Begin VB.Label Label1 
      Caption         =   "Original properties:"
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   2
      Left            =   9100
      TabIndex        =   63
      Top             =   180
      Width           =   2208
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Style3DEffect:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   11
      Left            =   7428
      TabIndex        =   110
      Top             =   7980
      Width           =   1104
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   29
      Left            =   6120
      TabIndex        =   56
      Top             =   7800
      Width           =   552
      _ExtentX        =   974
      _ExtentY        =   974
      BorderColor     =   56
      Shape           =   29
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   24
      Left            =   5064
      TabIndex        =   55
      Top             =   7800
      Width           =   504
      _ExtentX        =   910
      _ExtentY        =   974
      BorderColor     =   55
      Shape           =   28
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   23
      Left            =   3780
      TabIndex        =   54
      Top             =   7800
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      BorderColor     =   54
      Shape           =   27
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   22
      Left            =   2700
      TabIndex        =   53
      Top             =   7800
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      BorderColor     =   53
      Shape           =   26
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   21
      Left            =   792
      TabIndex        =   51
      Top             =   7800
      Width           =   396
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   24
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   20
      Left            =   1872
      TabIndex        =   52
      Top             =   7800
      Width           =   396
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   25
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   19
      Left            =   6190
      TabIndex        =   44
      Top             =   6504
      Width           =   400
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   23
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   18
      Left            =   5085
      TabIndex        =   43
      Top             =   6504
      Width           =   450
      _ExtentX        =   804
      _ExtentY        =   974
      Shape           =   22
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   26
      Left            =   2700
      TabIndex        =   41
      Top             =   6504
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   20
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   25
      Left            =   3930
      TabIndex        =   42
      Top             =   6504
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   21
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   17
      Left            =   1770
      TabIndex        =   40
      Top             =   6504
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   19
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   16
      Left            =   540
      TabIndex        =   39
      Top             =   6504
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   18
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   15
      Left            =   2700
      TabIndex        =   29
      Top             =   5208
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   14
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   14
      Left            =   3780
      TabIndex        =   30
      Top             =   5208
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   15
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   13
      Left            =   4860
      TabIndex        =   31
      Top             =   5208
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   16
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   12
      Left            =   5940
      TabIndex        =   32
      Top             =   5208
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   17
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   562
      Index           =   5
      Left            =   5940
      TabIndex        =   13
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   995
      Shape           =   5
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   4
      Left            =   4860
      TabIndex        =   12
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   4
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   562
      Index           =   3
      Left            =   3780
      TabIndex        =   11
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   995
      Shape           =   3
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   2
      Left            =   2700
      TabIndex        =   10
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   2
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   1
      Left            =   1620
      TabIndex        =   9
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   1
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   2556
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      BackColor       =   16744576
      BackStyle       =   1
      FillColor       =   12648447
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   7
      Left            =   2700
      TabIndex        =   17
      Top             =   3912
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   8
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   8
      Left            =   1620
      TabIndex        =   16
      Top             =   3912
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   7
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   28
      Left            =   590
      TabIndex        =   27
      Top             =   5208
      Width           =   800
      _ExtentX        =   1439
      _ExtentY        =   974
      Shape           =   12
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   27
      Left            =   1620
      TabIndex        =   28
      Top             =   5208
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   13
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   11
      Left            =   6090
      TabIndex        =   20
      Top             =   3912
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   11
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   10
      Left            =   4860
      TabIndex        =   19
      Top             =   3912
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   10
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   9
      Left            =   540
      TabIndex        =   15
      Top             =   3912
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   6
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   6
      Left            =   3780
      TabIndex        =   18
      Top             =   3912
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   9
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Style3D:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   10
      Left            =   7428
      TabIndex        =   108
      Top             =   7560
      Width           =   648
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BorderWidth:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   9
      Left            =   7404
      TabIndex        =   74
      Top             =   2280
      Width           =   1056
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BorderStyle:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   8
      Left            =   7404
      TabIndex        =   72
      Top             =   1860
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "It applies to some shapes."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Index           =   10
      Left            =   12532
      TabIndex        =   93
      Top             =   4980
      Width           =   2412
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pie"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   29
      Left            =   5880
      TabIndex        =   62
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "diamond"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   28
      Left            =   480
      TabIndex        =   33
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "trapezoid"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   27
      Left            =   1560
      TabIndex        =   34
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "arrow"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   26
      Left            =   2640
      TabIndex        =   47
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "crescent"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   25
      Left            =   3720
      TabIndex        =   48
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "shield"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   24
      Left            =   4800
      TabIndex        =   61
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "talk"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   23
      Left            =   3720
      TabIndex        =   60
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CurvingFactor:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   6
      Left            =   7404
      TabIndex        =   90
      Top             =   5040
      Width           =   1128
   End
   Begin VB.Label lblCurvingFactorValue 
      Caption         =   "0"
      Height          =   300
      Left            =   12100
      TabIndex        =   92
      Top             =   4980
      Width           =   624
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cloud"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   22
      Left            =   2640
      TabIndex        =   59
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "location"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   21
      Left            =   480
      TabIndex        =   57
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "speaker"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   20
      Left            =   1560
      TabIndex        =   58
      Top             =   8496
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "egg"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   19
      Left            =   5880
      TabIndex        =   50
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "drop"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   18
      Left            =   4800
      TabIndex        =   49
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vertices:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   5
      Left            =   7428
      TabIndex        =   114
      Top             =   8820
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Additional value used by these shapes: regular polygon, star and jagged star."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   9868
      TabIndex        =   116
      Top             =   8852
      Width           =   5724
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTest.frx":0024
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   7
      Left            =   9868
      TabIndex        =   107
      Top             =   7056
      Width           =   6680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shift:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   7428
      TabIndex        =   105
      Top             =   7140
      Width           =   396
   End
   Begin VB.Label Label1 
      Caption         =   "High is antialiased, Low is like the old Shape control. The default is High."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   12100
      TabIndex        =   101
      Top             =   6180
      Width           =   2916
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RotationDegrees:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   7428
      TabIndex        =   102
      Top             =   6720
      Width           =   1380
   End
   Begin VB.Label lblRotationDegreesValue 
      Caption         =   "0"
      Height          =   300
      Left            =   12088
      TabIndex        =   117
      Top             =   6752
      Width           =   408
   End
   Begin VB.Label lblOpacityValue 
      Caption         =   "100"
      Height          =   300
      Left            =   12088
      TabIndex        =   98
      Top             =   5916
      Width           =   624
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Opacity:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   7404
      TabIndex        =   96
      Top             =   5880
      Width           =   648
   End
   Begin VB.Label Label1 
      Caption         =   "It determines if it will produce mouse events."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   12160
      TabIndex        =   89
      Top             =   4560
      Width           =   4272
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "heart"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   17
      Left            =   1560
      TabIndex        =   46
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "jagged star"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   16
      Left            =   480
      TabIndex        =   45
      Top             =   7204
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "parallelogram"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   15
      Left            =   2640
      TabIndex        =   35
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "semicircle"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   14
      Left            =   3720
      TabIndex        =   36
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "regular polygon"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   13
      Left            =   4800
      TabIndex        =   37
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "star"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   12
      Left            =   5880
      TabIndex        =   38
      Top             =   5908
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "kite"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   11
      Left            =   5880
      TabIndex        =   26
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rhombus"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   10
      Left            =   4800
      TabIndex        =   25
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "New shapes:"
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   1
      Left            =   504
      TabIndex        =   14
      Top             =   3492
      Width           =   2208
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle equilateral"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   9
      Left            =   480
      TabIndex        =   21
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle isosceles"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   8
      Left            =   1560
      TabIndex        =   22
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle scalene"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   7
      Left            =   2640
      TabIndex        =   23
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle right"
      ForeColor       =   &H80000011&
      Height          =   492
      Index           =   6
      Left            =   3720
      TabIndex        =   24
      Top             =   4612
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rounded square"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   5
      Left            =   5880
      TabIndex        =   6
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rounded rectangle"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   4
      Left            =   4800
      TabIndex        =   5
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "circle"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "oval"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "square"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rectangle"
      ForeColor       =   &H80000011&
      Height          =   480
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1300
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "New ShapeEx control:"
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   0
      Left            =   504
      TabIndex        =   7
      Top             =   2136
      Width           =   2208
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBackStyle_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).BackStyle = Val(cboBackStyle.ListIndex)
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).BackStyle = Val(cboBackStyle.ListIndex)
    Next
End Sub

Private Sub cboBorderStyle_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).BorderStyle = Val(cboBorderStyle.ListIndex)
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).BorderStyle = Val(cboBorderStyle.ListIndex)
    Next
End Sub

Private Sub cboClickable_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Clickable = cboClickable.ListIndex = 1
    Next
End Sub

Private Sub cboFillStyle_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).FillStyle = Val(cboFillStyle.ListIndex)
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).FillStyle = Val(cboFillStyle.ListIndex)
    Next
End Sub

Private Sub cboFlipped_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Flipped = cboFlipped.ListIndex
    Next
End Sub

Private Sub cboQuality_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Quality = cboQuality.ListIndex
    Next
End Sub

Private Sub cboStyle3D_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Style3D = cboStyle3D.ListIndex
    Next
End Sub

Private Sub cboStyle3DEffect_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Style3DEffect = cboStyle3DEffect.ListIndex
    Next
End Sub

Private Sub chkAutoRotation_Click()
    tmrAutoRotation.Enabled = (chkAutoRotation.Value = 1)
End Sub

Private Sub cmdChangeBackColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        shpBackColor.BackColor = iDlg.Color
        SetBackColor
    End If
End Sub

Private Sub SetBackColor()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).BackColor = shpBackColor.BackColor
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).BackColor = shpBackColor.BackColor
    Next
End Sub

Private Sub cmdChangeBorderColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        shpBorderColor.BackColor = iDlg.Color
        SetBorderColor
    End If
End Sub

Private Sub SetBorderColor()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).BorderColor = shpBorderColor.BackColor
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).BorderColor = shpBorderColor.BackColor
    Next
End Sub

Private Sub cmdChangeFillColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        shpFillColor.BackColor = iDlg.Color
        SetFillColor
    End If
End Sub

Private Sub SetFillColor()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).FillColor = shpFillColor.BackColor
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).FillColor = shpFillColor.BackColor
    Next
End Sub

Private Sub cmdDecreaseBordderWidth_Click()
    If Val(txtBorderWidth.Text) > 1 Then
        txtBorderWidth.Text = Val(txtBorderWidth.Text) - 1
    End If
End Sub

Private Sub cmdIncreaseBordderWidth_Click()
    txtBorderWidth.Text = Val(txtBorderWidth.Text) + 1
End Sub

Private Sub Form_Load()
    LoadCombos
    cboBorderStyle.ListIndex = ShapeEx1(0).BorderStyle
    cboStyle3D.ListIndex = ShapeEx1(0).Style3D
    cboStyle3DEffect.ListIndex = ShapeEx1(0).Style3DEffect
    cboUseSubclassing.ListIndex = ShapeEx1(0).UseSubclassing
    txtShift_Change
    cboBackStyle.ListIndex = ShapeEx1(0).BackStyle
    shpBackColor.BackColor = ShapeEx1(0).BackColor
    SetBackColor
    shpBorderColor.BackColor = ShapeEx1(0).BorderColor
    SetBorderColor
    txtBorderWidth.Text = ShapeEx1(0).BorderWidth
    shpFillColor.BackColor = ShapeEx1(0).FillColor
    SetFillColor
    cboFillStyle.ListIndex = ShapeEx1(0).FillStyle
    cboFlipped.ListIndex = ShapeEx1(0).Flipped
    cboQuality.ListIndex = ShapeEx1(0).Quality
    cboClickable.ListIndex = CLng(ShapeEx1(0).Clickable) * -1
End Sub

Private Sub LoadCombos()
    cboBorderStyle.Clear
    cboBorderStyle.AddItem "vbTransparent"
    cboBorderStyle.AddItem "vbBSSolid"
    cboBorderStyle.AddItem "vbBSDash"
    cboBorderStyle.AddItem "vbBSDot"
    cboBorderStyle.AddItem "vbBSDashDot"
    cboBorderStyle.AddItem "vbBSDashDotDot"
    cboBorderStyle.AddItem "vbBSInsideSolid"
    
    cboStyle3D.Clear
    cboStyle3D.AddItem "seStyle3DNone"
    cboStyle3D.AddItem "seStyle3DLight"
    cboStyle3D.AddItem "seStyle3DShadow"
    cboStyle3D.AddItem "seStyle3DBoth"
    
    cboStyle3DEffect.Clear
    cboStyle3DEffect.AddItem "seStyle3EffectAuto"
    cboStyle3DEffect.AddItem "seStyle3EffectDifusse"
    cboStyle3DEffect.AddItem "seStyle3EffectGem"

    cboUseSubclassing.Clear
    cboUseSubclassing.AddItem "seSCNo"
    cboUseSubclassing.AddItem "seSCYes"
    cboUseSubclassing.AddItem "seSCNotInIDE"
    cboUseSubclassing.AddItem "seSCNotInIDEDesignTime"
    
    cboBackStyle.Clear
    cboBackStyle.AddItem "seTransparent"
    cboBackStyle.AddItem "seOpaque"
    
    cboFillStyle.Clear
    cboFillStyle.AddItem "seFSSolid"
    cboFillStyle.AddItem "seFSTransparent"
    
    cboFlipped.Clear
    cboFlipped.AddItem "seFlippedNo"
    cboFlipped.AddItem "seFlippedHorizontally"
    cboFlipped.AddItem "seFlippedVertically"
    cboFlipped.AddItem "seFlippedBoth"
    
    cboQuality.Clear
    cboQuality.AddItem "seQualityLow"
    cboQuality.AddItem "seQualityHigh"
    
    cboClickable.Clear
    cboClickable.AddItem "False"
    cboClickable.AddItem "True"
End Sub

Private Sub ShapeEx1_Click(Index As Integer)
    MsgBox "Control is clickable"
End Sub

Private Sub shpBackColor_Click()
    cmdChangeBackColor_Click
End Sub

Private Sub shpBorderColor_Click()
    cmdChangeBorderColor_Click
End Sub

Private Sub shpFillColor_Click()
    cmdChangeFillColor_Click
End Sub

Private Sub sldOpacity_Change()
    sldOpacity_Click
End Sub

Private Sub sldOpacity_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Opacity = sldOpacity.Value
    Next
    lblOpacityValue.Caption = sldOpacity.Value
End Sub

Private Sub sldOpacity_Scroll()
    sldOpacity_Change
End Sub

Private Sub sldRotationDegrees_Change()
    sldRotationDegrees_Click
    tmrRefresh.Enabled = True
End Sub

Private Sub sldRotationDegrees_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).RotationDegrees = sldRotationDegrees.Value
    Next
    lblRotationDegreesValue.Caption = sldRotationDegrees.Value
End Sub

Private Sub sldRotationDegrees_Scroll()
    sldRotationDegrees_Change
End Sub

Private Sub txtBorderWidth_Change()
    Dim c As Long
    
    If Not IsNumeric(txtBorderWidth.Text) And (Trim(txtBorderWidth.Text) <> "") And (Trim(txtBorderWidth.Text) <> "-") Then
        MsgBox "Value must be numeric", vbExclamation
        txtBorderWidth.Text = ""
        Exit Sub
    End If
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).BorderWidth = Val(txtBorderWidth.Text)
    Next
    For c = Shape1.LBound To Shape1.UBound
        Shape1(c).BorderWidth = Val(txtBorderWidth.Text)
    Next
End Sub

Private Sub tmrAutoRotation_Timer()
    If sldRotationDegrees.Value < 3 Then sldRotationDegrees.Value = 360
    sldRotationDegrees.Value = sldRotationDegrees.Value - 3
    sldRotationDegrees.Value = ShapeEx1(0).RotationDegrees
End Sub

Private Sub tmrRefresh_Timer()
    Me.Refresh
    tmrRefresh.Enabled = False
End Sub

Private Sub txtShift_Change()
    Dim c As Long
    
    If Not IsNumeric(txtShift.Text) And (Trim(txtShift.Text) <> "") And (Trim(txtShift.Text) <> "-") Then
        MsgBox "Value must be numeric", vbExclamation
        txtShift.Text = ""
        Exit Sub
    End If
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Shift = Val(txtShift.Text)
    Next
End Sub

Private Sub txtVertices_Change()
    Dim c As Long
    
    If Not IsNumeric(txtVertices.Text) And (Trim(txtVertices.Text) <> "") Then
        MsgBox "Value must be numeric", vbExclamation
        txtVertices.Text = ""
        Exit Sub
    End If
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).Vertices = Val(txtVertices.Text)
    Next
End Sub

Private Sub sldCurvingFactor_Change()
    sldCurvingFactor_Click
    tmrRefresh.Enabled = True
End Sub

Private Sub sldCurvingFactor_Click()
    Dim c As Long
    
    For c = ShapeEx1.LBound To ShapeEx1.UBound
        ShapeEx1(c).CurvingFactor = sldCurvingFactor.Value
    Next
    lblCurvingFactorValue.Caption = sldCurvingFactor.Value
End Sub

Private Sub sldCurvingFactor_Scroll()
    sldCurvingFactor_Change
End Sub

