VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Shapes"
   ClientHeight    =   10512
   ClientLeft      =   1692
   ClientTop       =   1128
   ClientWidth     =   9336
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10512
   ScaleWidth      =   9336
   Begin VB.CheckBox chkMirrored 
      Caption         =   "Mirrored"
      Height          =   264
      Left            =   3816
      TabIndex        =   59
      Top             =   10152
      Width           =   1632
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7272
      Top             =   7416
   End
   Begin VB.CheckBox chkAutoRotation 
      Caption         =   "Animate"
      Height          =   264
      Left            =   6348
      TabIndex        =   39
      Top             =   7812
      Width           =   968
   End
   Begin VB.Timer tmrAutoRotation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6876
      Top             =   7416
   End
   Begin VB.TextBox txtVertices 
      Alignment       =   1  'Right Justify
      Height          =   324
      Left            =   1464
      TabIndex        =   36
      Text            =   "5"
      Top             =   9684
      Width           =   624
   End
   Begin VB.TextBox txtShift 
      Alignment       =   1  'Right Justify
      Height          =   324
      Left            =   1464
      TabIndex        =   35
      Text            =   "20"
      Top             =   9256
      Width           =   624
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Left            =   2916
      TabIndex        =   25
      Top             =   7416
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
      Left            =   2916
      TabIndex        =   27
      Top             =   7776
      Width           =   2904
      _ExtentX        =   5122
      _ExtentY        =   529
      _Version        =   327682
      Max             =   360
   End
   Begin ComctlLib.Slider sldCurvingFactor 
      Height          =   300
      Left            =   2916
      TabIndex        =   45
      Top             =   8136
      Width           =   2904
      _ExtentX        =   5122
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   100
   End
   Begin VB.Label Label1 
      Caption         =   "Flips the shape horizontally"
      Height          =   300
      Index           =   11
      Left            =   1512
      TabIndex        =   58
      Top             =   10152
      Width           =   2208
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mirrored:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   7
      Left            =   504
      TabIndex        =   57
      Top             =   10152
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "It applies to shapes that are not curved"
      Height          =   408
      Index           =   10
      Left            =   6176
      TabIndex        =   56
      Top             =   8172
      Width           =   3792
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   29
      Left            =   8100
      Top             =   5328
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      FillColor       =   16744576
      FillStyle       =   0
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      Height          =   492
      Index           =   29
      Left            =   8040
      TabIndex        =   55
      Top             =   6028
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "diamond"
      Height          =   492
      Index           =   28
      Left            =   6960
      TabIndex        =   54
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   28
      Left            =   7070
      Top             =   2736
      Width           =   800
      _ExtentX        =   1418
      _ExtentY        =   974
      Shape           =   12
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "trapezoid"
      Height          =   492
      Index           =   27
      Left            =   8040
      TabIndex        =   53
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   27
      Left            =   8100
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   13
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "arrow"
      Height          =   492
      Index           =   26
      Left            =   6960
      TabIndex        =   52
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   26
      Left            =   7020
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   20
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "crescent"
      Height          =   492
      Index           =   25
      Left            =   8040
      TabIndex        =   51
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   25
      Left            =   8250
      Top             =   4032
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   21
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   24
      Left            =   7220
      Top             =   5328
      Width           =   500
      _ExtentX        =   889
      _ExtentY        =   974
      Shape           =   28
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "shield"
      Height          =   492
      Index           =   24
      Left            =   6960
      TabIndex        =   50
      Top             =   6028
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   23
      Left            =   5940
      Top             =   5328
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   27
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "talk"
      Height          =   492
      Index           =   23
      Left            =   5880
      TabIndex        =   49
      Top             =   6028
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CurvingFactor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   6
      Left            =   504
      TabIndex        =   48
      Top             =   8136
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "-100-100"
      Height          =   300
      Index           =   9
      Left            =   2052
      TabIndex        =   47
      Top             =   8100
      Width           =   768
   End
   Begin VB.Label lblCurvingFactorValue 
      Caption         =   "0"
      Height          =   300
      Left            =   5900
      TabIndex        =   46
      Top             =   8172
      Width           =   624
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cloud"
      Height          =   492
      Index           =   22
      Left            =   4800
      TabIndex        =   44
      Top             =   6028
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   22
      Left            =   4860
      Top             =   5328
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   26
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   21
      Left            =   2950
      Top             =   5328
      Width           =   400
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   24
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "location"
      Height          =   492
      Index           =   21
      Left            =   2640
      TabIndex        =   43
      Top             =   6028
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   20
      Left            =   4030
      Top             =   5328
      Width           =   400
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   25
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "speaker"
      Height          =   492
      Index           =   20
      Left            =   3720
      TabIndex        =   42
      Top             =   6028
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "egg"
      Height          =   492
      Index           =   19
      Left            =   1560
      TabIndex        =   41
      Top             =   6028
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   19
      Left            =   1870
      Top             =   5328
      Width           =   400
      _ExtentX        =   699
      _ExtentY        =   974
      Shape           =   23
      FillColor       =   16744576
      FillStyle       =   0
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "drop"
      Height          =   492
      Index           =   18
      Left            =   480
      TabIndex        =   40
      Top             =   6028
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   18
      Left            =   765
      Top             =   5328
      Width           =   450
      _ExtentX        =   804
      _ExtentY        =   974
      Shape           =   22
      FillColor       =   16744576
      FillStyle       =   0
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vertices:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   5
      Left            =   504
      TabIndex        =   38
      Top             =   9720
      Width           =   804
   End
   Begin VB.Label Label1 
      Caption         =   "Additional value used by these shapes: regular polygon and star."
      Height          =   372
      Index           =   8
      Left            =   2172
      TabIndex        =   37
      Top             =   9756
      Width           =   5124
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTest.frx":0000
      Height          =   444
      Index           =   7
      Left            =   2172
      TabIndex        =   34
      Top             =   9216
      Width           =   6900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shift:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   4
      Left            =   504
      TabIndex        =   33
      Top             =   9300
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "High is antialiased, Low is like the old Shape control. The default is High"
      Height          =   300
      Index           =   6
      Left            =   1512
      TabIndex        =   32
      Top             =   8856
      Width           =   5628
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Quality:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   504
      TabIndex        =   31
      Top             =   8892
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "0-360"
      Height          =   300
      Index           =   5
      Left            =   2196
      TabIndex        =   30
      Top             =   7776
      Width           =   624
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RotationDegrees:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   2
      Left            =   504
      TabIndex        =   29
      Top             =   7776
      Width           =   1572
   End
   Begin VB.Label lblRotationDegreesValue 
      Caption         =   "0"
      Height          =   300
      Left            =   5900
      TabIndex        =   28
      Top             =   7812
      Width           =   408
   End
   Begin VB.Label lblOpacityValue 
      Caption         =   "100"
      Height          =   300
      Left            =   5900
      TabIndex        =   26
      Top             =   7452
      Width           =   624
   End
   Begin VB.Label Label1 
      Caption         =   "0-100"
      Height          =   300
      Index           =   4
      Left            =   2196
      TabIndex        =   24
      Top             =   7380
      Width           =   624
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Opacity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   1
      Left            =   504
      TabIndex        =   23
      Top             =   7416
      Width           =   756
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clickable:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   504
      TabIndex        =   22
      Top             =   8532
      Width           =   888
   End
   Begin VB.Label Label1 
      Caption         =   "It determines whether it will raise mouse events"
      Height          =   300
      Index           =   3
      Left            =   1512
      TabIndex        =   21
      Top             =   8532
      Width           =   4152
   End
   Begin VB.Label Label1 
      Caption         =   "New properties:"
      Height          =   300
      Index           =   2
      Left            =   504
      TabIndex        =   20
      Top             =   6912
      Width           =   2208
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   17
      Left            =   6090
      Top             =   4032
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   19
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "heart"
      Height          =   492
      Index           =   17
      Left            =   5880
      TabIndex        =   19
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   16
      Left            =   4860
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   18
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "jagged star"
      Height          =   492
      Index           =   16
      Left            =   4800
      TabIndex        =   18
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   15
      Left            =   540
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   14
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "parallelogram"
      Height          =   492
      Index           =   15
      Left            =   480
      TabIndex        =   17
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   14
      Left            =   1620
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   15
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "semicircle"
      Height          =   492
      Index           =   14
      Left            =   1560
      TabIndex        =   16
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   13
      Left            =   2700
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   16
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "regular polygon"
      Height          =   492
      Index           =   13
      Left            =   2640
      TabIndex        =   15
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   12
      Left            =   3780
      Top             =   4032
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   17
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "star"
      Height          =   492
      Index           =   12
      Left            =   3720
      TabIndex        =   14
      Top             =   4732
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   11
      Left            =   6090
      Top             =   2736
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   974
      Shape           =   11
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "kite"
      Height          =   492
      Index           =   11
      Left            =   5880
      TabIndex        =   13
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   10
      Left            =   4860
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   10
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rhombus"
      Height          =   492
      Index           =   10
      Left            =   4800
      TabIndex        =   12
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   9
      Left            =   540
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   6
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "New shapes:"
      Height          =   300
      Index           =   1
      Left            =   504
      TabIndex        =   11
      Top             =   2196
      Width           =   2208
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle equilateral"
      Height          =   492
      Index           =   9
      Left            =   480
      TabIndex        =   10
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   8
      Left            =   1620
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   7
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle isosceles"
      Height          =   492
      Index           =   8
      Left            =   1560
      TabIndex        =   9
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   7
      Left            =   2700
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   8
      FillColor       =   16744576
      FillStyle       =   0
      Shift           =   20
      ShiftPutAutomatically=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle scalene"
      Height          =   492
      Index           =   7
      Left            =   2640
      TabIndex        =   8
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   6
      Left            =   3780
      Top             =   2736
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   9
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "triangle right"
      Height          =   492
      Index           =   6
      Left            =   3720
      TabIndex        =   7
      Top             =   3436
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   562
      Index           =   5
      Left            =   5940
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   995
      Shape           =   5
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rounded square"
      Height          =   480
      Index           =   5
      Left            =   5880
      TabIndex        =   6
      Top             =   1420
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   4
      Left            =   4860
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   4
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rounded rectangle"
      Height          =   480
      Index           =   4
      Left            =   4800
      TabIndex        =   5
      Top             =   1420
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "circle"
      Height          =   480
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Top             =   1420
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   562
      Index           =   3
      Left            =   3780
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   995
      Shape           =   3
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "oval"
      Height          =   480
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   1420
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   2
      Left            =   2700
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   2
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "square"
      Height          =   480
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1420
      Width           =   1020
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   1
      Left            =   1620
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      Shape           =   1
      FillColor       =   16744576
      FillStyle       =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rectangle"
      Height          =   480
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1420
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Original shapes:"
      Height          =   300
      Index           =   0
      Left            =   504
      TabIndex        =   0
      Top             =   180
      Width           =   2208
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   552
      Index           =   0
      Left            =   540
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   974
      FillColor       =   16744576
      FillStyle       =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutoRotation_Click()
    tmrAutoRotation.Enabled = (chkAutoRotation.Value = 1)
End Sub

Private Sub chkMirrored_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.Mirrored = chkMirrored.Value = 1
        End If
    Next
End Sub

Private Sub sldOpacity_Change()
    sldOpacity_Click
End Sub

Private Sub sldOpacity_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.Opacity = sldOpacity.Value
        End If
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
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.RotationDegrees = sldRotationDegrees.Value
        End If
    Next
    lblRotationDegreesValue.Caption = sldRotationDegrees.Value
End Sub

Private Sub sldRotationDegrees_Scroll()
    sldRotationDegrees_Change
End Sub

Private Sub tmrAutoRotation_Timer()
    sldRotationDegrees.Value = sldRotationDegrees.Value + 1
    sldRotationDegrees.Value = ShapeEx1(0).RotationDegrees
End Sub

Private Sub tmrRefresh_Timer()
    Me.Refresh
    tmrRefresh.Enabled = False
End Sub

Private Sub txtShift_Change()
    If Not IsNumeric(txtShift.Text) And (Trim(txtShift.Text) <> "") And (Trim(txtShift.Text) <> "-") Then
        MsgBox "Value must be numeric", vbExclamation
        txtShift.Text = ""
        Exit Sub
    End If
    
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.Shift = Val(txtShift.Text)
        End If
    Next
End Sub

Private Sub txtVertices_Change()
    If Not IsNumeric(txtVertices.Text) And (Trim(txtVertices.Text) <> "") Then
        MsgBox "Value must be numeric", vbExclamation
        txtVertices.Text = ""
        Exit Sub
    End If
    
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.Vertices = Val(txtVertices.Text)
        End If
    Next
End Sub

Private Sub sldCurvingFactor_Change()
    sldCurvingFactor_Click
    tmrRefresh.Enabled = True
End Sub

Private Sub sldCurvingFactor_Click()
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) = "ShapeEx" Then
            iCtl.CurvingFactor = sldCurvingFactor.Value
        End If
    Next
    lblCurvingFactorValue.Caption = sldCurvingFactor.Value
End Sub

Private Sub sldCurvingFactor_Scroll()
    sldCurvingFactor_Change
End Sub

