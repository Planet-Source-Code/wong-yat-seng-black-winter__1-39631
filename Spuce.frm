VERSION 5.00
Begin VB.Form BlackWinter 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Black Winter"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6495
   ForeColor       =   &H00000000&
   Icon            =   "Spuce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer EBull 
      Left            =   3360
      Top             =   4920
   End
   Begin VB.Timer StatusMove 
      Left            =   3360
      Top             =   4800
   End
   Begin VB.Timer ReloadMove 
      Left            =   3360
      Top             =   4680
   End
   Begin VB.Timer Protection 
      Left            =   3360
      Top             =   4560
   End
   Begin VB.Timer ExploMove 
      Left            =   3360
      Top             =   4440
   End
   Begin VB.Timer RicoMove 
      Left            =   3360
      Top             =   4320
   End
   Begin VB.Timer GalagaMove 
      Left            =   3360
      Top             =   4200
   End
   Begin VB.Timer NukeTimer 
      Left            =   3360
      Top             =   4080
   End
   Begin VB.Timer EnemyMove 
      Left            =   3360
      Top             =   3960
   End
   Begin VB.Timer DistMove 
      Left            =   3360
      Top             =   3840
   End
   Begin VB.Timer FuelMove 
      Left            =   3360
      Top             =   3720
   End
   Begin VB.Timer BulletMove 
      Left            =   3360
      Top             =   3600
   End
   Begin VB.Timer ShowStats 
      Left            =   3360
      Top             =   3480
   End
   Begin VB.Timer ShipMove 
      Left            =   3360
      Top             =   3360
   End
   Begin VB.Image PWup 
      Height          =   150
      Index           =   4
      Left            =   360
      Picture         =   "Spuce.frx":0442
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image PWup 
      Height          =   240
      Index           =   3
      Left            =   120
      Picture         =   "Spuce.frx":0884
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image PWup 
      Height          =   195
      Index           =   2
      Left            =   600
      Picture         =   "Spuce.frx":0CC6
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   180
   End
   Begin VB.Image PWup 
      Height          =   255
      Index           =   1
      Left            =   360
      Picture         =   "Spuce.frx":1780
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image PWup 
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "Spuce.frx":1BC2
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   20
      Left            =   720
      Picture         =   "Spuce.frx":2004
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   19
      Left            =   600
      Picture         =   "Spuce.frx":2446
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   18
      Left            =   480
      Picture         =   "Spuce.frx":2888
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   17
      Left            =   720
      Picture         =   "Spuce.frx":2CCA
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   16
      Left            =   720
      Picture         =   "Spuce.frx":310C
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   15
      Left            =   720
      Picture         =   "Spuce.frx":354E
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   14
      Left            =   720
      Picture         =   "Spuce.frx":3990
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   13
      Left            =   720
      Picture         =   "Spuce.frx":3DD2
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   12
      Left            =   720
      Picture         =   "Spuce.frx":4214
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   11
      Left            =   600
      Picture         =   "Spuce.frx":4656
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   10
      Left            =   600
      Picture         =   "Spuce.frx":4A98
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   9
      Left            =   600
      Picture         =   "Spuce.frx":4EDA
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   8
      Left            =   600
      Picture         =   "Spuce.frx":531C
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   7
      Left            =   600
      Picture         =   "Spuce.frx":575E
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   6
      Left            =   600
      Picture         =   "Spuce.frx":5BA0
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   5
      Left            =   480
      Picture         =   "Spuce.frx":5FE2
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   4
      Left            =   480
      Picture         =   "Spuce.frx":6424
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   3
      Left            =   480
      Picture         =   "Spuce.frx":6866
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   2
      Left            =   480
      Picture         =   "Spuce.frx":6CA8
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   1
      Left            =   480
      Picture         =   "Spuce.frx":70EA
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label rtb1 
      BackColor       =   &H00000000&
      Caption         =   $"Spuce.frx":752C
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3135
      Left            =   3240
      TabIndex        =   70
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   31
      Left            =   1200
      Picture         =   "Spuce.frx":77FB
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   49
      Left            =   5640
      Picture         =   "Spuce.frx":9F9D
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   48
      Left            =   1080
      Picture         =   "Spuce.frx":B20F
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   47
      Left            =   0
      Picture         =   "Spuce.frx":C481
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   465
      Index           =   45
      Left            =   3000
      Picture         =   "Spuce.frx":D6F3
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   44
      Left            =   3720
      Picture         =   "Spuce.frx":E965
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   42
      Left            =   2400
      Picture         =   "Spuce.frx":FBD7
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   43
      Left            =   480
      Picture         =   "Spuce.frx":10E49
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   465
      Index           =   46
      Left            =   4920
      Picture         =   "Spuce.frx":120BB
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   41
      Left            =   4320
      Picture         =   "Spuce.frx":1332D
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Galaga 
      Height          =   480
      Index           =   40
      Left            =   1680
      Picture         =   "Spuce.frx":1459F
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Index           =   3
      Left            =   960
      TabIndex        =   68
      Top             =   960
      Width           =   1980
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Index           =   2
      Left            =   960
      TabIndex        =   67
      Top             =   800
      Width           =   1980
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Index           =   1
      Left            =   960
      TabIndex        =   66
      Top             =   640
      Width           =   1980
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   165
      Index           =   0
      Left            =   960
      TabIndex        =   65
      Top             =   480
      Width           =   1980
   End
   Begin VB.Shape Prot_Circle 
      BorderColor     =   &H00FFFFC0&
      Height          =   375
      Left            =   480
      Shape           =   2  'Oval
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label GunLock_R 
      BackColor       =   &H00C0C0FF&
      Height          =   75
      Left            =   3120
      TabIndex        =   64
      Top             =   2930
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label GunLock_L 
      BackColor       =   &H00C0C0FF&
      Height          =   80
      Left            =   2640
      TabIndex        =   63
      Top             =   2930
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblGun_Lock 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "GUN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EFF9BB&
      Height          =   165
      Left            =   2760
      TabIndex        =   62
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image picExplo4 
      Height          =   300
      Left            =   120
      Picture         =   "Spuce.frx":15811
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image picExplo3 
      Height          =   300
      Left            =   120
      Picture         =   "Spuce.frx":16CAB
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image picExplo2 
      Height          =   300
      Left            =   120
      Picture         =   "Spuce.frx":18145
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image PicExplo1 
      Height          =   300
      Left            =   120
      Picture         =   "Spuce.frx":195DF
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image picRico 
      Height          =   180
      Left            =   480
      Picture         =   "Spuce.frx":1AA79
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image eBullet 
      Height          =   105
      Index           =   0
      Left            =   480
      Picture         =   "Spuce.frx":1B4CB
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   300
      Picture         =   "Spuce.frx":1B90D
      Stretch         =   -1  'True
      Top             =   2700
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   570
      Picture         =   "Spuce.frx":1BD4F
      Stretch         =   -1  'True
      Top             =   2740
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   60
      Picture         =   "Spuce.frx":1C809
      Stretch         =   -1  'True
      Top             =   2710
      Width           =   255
   End
   Begin VB.Image Indic 
      Height          =   120
      Left            =   120
      Picture         =   "Spuce.frx":1CC4B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dist."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   960
      Width           =   615
   End
   Begin VB.Label picDist 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   120
      TabIndex        =   39
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   9
      Left            =   600
      TabIndex        =   38
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   8
      Left            =   600
      TabIndex        =   37
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   7
      Left            =   600
      TabIndex        =   36
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   6
      Left            =   600
      TabIndex        =   35
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   5
      Left            =   600
      TabIndex        =   34
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   4
      Left            =   600
      TabIndex        =   33
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   3
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   600
      TabIndex        =   31
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   600
      TabIndex        =   30
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Fuel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   600
      TabIndex        =   29
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   9
      Left            =   360
      TabIndex        =   28
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   8
      Left            =   360
      TabIndex        =   27
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   6
      Left            =   360
      TabIndex        =   25
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   4
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   3
      Left            =   360
      TabIndex        =   22
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Shield 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Ammo 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label picFuel 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label picAmmo 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label picShields 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Image picBomb 
      Height          =   240
      Left            =   120
      Picture         =   "Spuce.frx":1D08D
      Stretch         =   -1  'True
      Top             =   480
      Width           =   240
   End
   Begin VB.Image picLives 
      Height          =   240
      Left            =   120
      Picture         =   "Spuce.frx":1D4CF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Stats 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Stats 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label StatBack 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Boss_BARfore 
      BackColor       =   &H0080FFFF&
      Height          =   60
      Left            =   1540
      TabIndex        =   60
      Top             =   80
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Boss_BARback 
      BackColor       =   &H000000FF&
      Height          =   60
      Left            =   1540
      TabIndex        =   59
      Top             =   80
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblWeaponLaser 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "TYPHOON"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   165
      Index           =   4
      Left            =   885
      TabIndex        =   58
      Top             =   2895
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblWeaponLaser 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "WAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   165
      Index           =   3
      Left            =   885
      TabIndex        =   57
      Top             =   2895
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblWeaponLaser 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "NAPALM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   2
      Left            =   885
      TabIndex        =   56
      Top             =   2895
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblWeaponLaser 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "VULCAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   165
      Index           =   1
      Left            =   885
      TabIndex        =   55
      Top             =   2895
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblWeaponLaser 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "LASER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   0
      Left            =   885
      TabIndex        =   54
      Top             =   2895
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   38
      Left            =   960
      Picture         =   "Spuce.frx":1D911
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   37
      Left            =   960
      Picture         =   "Spuce.frx":200B3
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   36
      Left            =   960
      Picture         =   "Spuce.frx":22855
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   35
      Left            =   960
      Picture         =   "Spuce.frx":24FF7
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   34
      Left            =   960
      Picture         =   "Spuce.frx":27799
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   33
      Left            =   960
      Picture         =   "Spuce.frx":29F3B
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   32
      Left            =   960
      Picture         =   "Spuce.frx":2C6DD
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   30
      Left            =   1200
      Picture         =   "Spuce.frx":2EE7F
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   29
      Left            =   1200
      Picture         =   "Spuce.frx":31621
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   28
      Left            =   1200
      Picture         =   "Spuce.frx":33DC3
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   27
      Left            =   1200
      Picture         =   "Spuce.frx":36565
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   26
      Left            =   1200
      Picture         =   "Spuce.frx":38D07
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   25
      Left            =   1200
      Picture         =   "Spuce.frx":3B4A9
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   23
      Left            =   1440
      Picture         =   "Spuce.frx":3DC4B
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   22
      Left            =   1440
      Picture         =   "Spuce.frx":403ED
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   21
      Left            =   1440
      Picture         =   "Spuce.frx":42B8F
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   20
      Left            =   1440
      Picture         =   "Spuce.frx":45331
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   19
      Left            =   1440
      Picture         =   "Spuce.frx":47AD3
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   18
      Left            =   1440
      Picture         =   "Spuce.frx":4A275
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   17
      Left            =   1440
      Picture         =   "Spuce.frx":4CA17
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   15
      Left            =   1680
      Picture         =   "Spuce.frx":4F1B9
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   14
      Left            =   1680
      Picture         =   "Spuce.frx":5195B
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   13
      Left            =   1680
      Picture         =   "Spuce.frx":540FD
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   12
      Left            =   1680
      Picture         =   "Spuce.frx":5689F
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   9
      Left            =   1680
      Picture         =   "Spuce.frx":59041
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   8
      Left            =   1680
      Picture         =   "Spuce.frx":5B7E3
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   39
      Left            =   960
      Picture         =   "Spuce.frx":5DF85
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   24
      Left            =   1200
      Picture         =   "Spuce.frx":60727
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   16
      Left            =   1440
      Picture         =   "Spuce.frx":62EC9
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   11
      Left            =   1680
      Picture         =   "Spuce.frx":6566B
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   10
      Left            =   1680
      Picture         =   "Spuce.frx":67E0D
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   7
      Left            =   1920
      Picture         =   "Spuce.frx":6A5AF
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   6
      Left            =   1920
      Picture         =   "Spuce.frx":6B069
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   5
      Left            =   1920
      Picture         =   "Spuce.frx":6BB23
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   4
      Left            =   1920
      Picture         =   "Spuce.frx":6C5DD
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   3
      Left            =   1920
      Picture         =   "Spuce.frx":6D097
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   2
      Left            =   1920
      Picture         =   "Spuce.frx":6DB51
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   1
      Left            =   1920
      Picture         =   "Spuce.frx":6E60B
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblNukeFore 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   900
      TabIndex        =   53
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblNukeBack 
      BackColor       =   &H000000C0&
      Height          =   735
      Left            =   900
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image NukeEM 
      Height          =   150
      Left            =   860
      Picture         =   "Spuce.frx":6F0C5
      Stretch         =   -1  'True
      Top             =   1980
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Galaga 
      Height          =   240
      Index           =   0
      Left            =   1920
      Picture         =   "Spuce.frx":6F507
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMyname 
      BackColor       =   &H00404040&
      Caption         =   $"Spuce.frx":6FFC1
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   960
      TabIndex        =   51
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   9
      Left            =   2040
      TabIndex        =   50
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   8
      Left            =   1920
      TabIndex        =   49
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   7
      Left            =   1800
      TabIndex        =   48
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   6
      Left            =   1680
      TabIndex        =   47
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   5
      Left            =   1560
      TabIndex        =   46
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   4
      Left            =   1440
      TabIndex        =   45
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   3
      Left            =   1320
      TabIndex        =   44
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   2
      Left            =   1200
      TabIndex        =   43
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   1
      Left            =   1080
      TabIndex        =   42
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Bullet 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Index           =   0
      Left            =   960
      TabIndex        =   41
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Stats 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label picLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   880
      TabIndex        =   4
      Top             =   10
      Width           =   615
   End
   Begin VB.Label Stats 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label picScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   15
      Width           =   735
   End
   Begin VB.Image Ship 
      Height          =   360
      Left            =   120
      Picture         =   "Spuce.frx":7005D
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   3135
      Left            =   840
      TabIndex        =   69
      Top             =   0
      Width           =   2450
   End
   Begin VB.Menu mnGame 
      Caption         =   "&Game"
      Begin VB.Menu mnNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnPause 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnQuit 
         Caption         =   "&Quit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "&About"
      Begin VB.Menu mnHow 
         Caption         =   "&How to Play"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnAbout2 
         Caption         =   "&About Black Winter"
      End
   End
End
Attribute VB_Name = "BlackWinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName
Dim ANS As Variant

Dim Ship_Dir_N As Boolean     'Ship Values...
Dim Ship_Dir_S As Boolean
Dim Ship_Dir_E As Boolean
Dim Ship_Dir_W As Boolean
Dim Ship_Speed As Integer

Dim Ship_Ammo As Integer    '0-9
Dim Ship_Shield As Integer  '0-9
Dim Ship_Fuel As Integer    '0-9
Dim Ship_Protect As Boolean

Dim Ship_Stats(3) As Double '0-Level
                            '1-Score
                            '2-Lives
                            '3-Nukes
Dim Gun_Lock As Integer

Dim Bullet_Type As Integer  '0-4
Dim Bullet_Damage As Integer
Dim Bullet_Speed As Integer
Const NukeStop = 60         '1 = 0.25 secs

Dim Galaga_Dam_Inv As Integer   'Galaga's Attributes
Dim Galaga_Dam(49) As Integer
Dim Galaga_Move_Pat_Inv As Integer
Dim Galaga_Move_Pat(49) As Integer
Dim Galaga_Worth_Inv As Single
Dim Galaga_Worth(49) As Single
Dim Galaga_Position_Pat As Integer
Dim Galaga_Num As Integer
Dim Galaga_Start As Integer
Dim Galaga_Start_POSx As Integer
Dim Galaga_Start_POSy As Integer
Dim Galaga_Speed As Integer
Dim Galaga_Speedx(49) As Integer
Dim Galaga_Init_Life As Double
Dim Galaga_Life(49) As Double
Dim Galaga_Boss_Bar As Double

Dim Bonus_Chance As Integer
Dim BossAppear As Boolean

Dim BulletTemp(4) As Integer
Dim TimeTemp(9) As Integer      'time Storage
Dim SmallTemp(3) As Integer  'Small items

Dim i As Integer
Dim x As Integer
Dim e As Integer        'for explosion
Dim b As Integer        'for bullets onli
Dim n As Integer        'for nuke counter
Dim g As Integer        'for galagas


Private Sub EBull_Timer()
'Create Enemy Bullets
For x = 0 To 49
    If Galaga(x).Visible = True Then
        If Galaga_Dam(x) >= (Rnd * 1000) Then
            For i = 0 To 20
                If eBullet(i).Visible = False Then
                    eBullet(i).Left = Galaga(x).Left + (Galaga(x).Width / 2)
                    eBullet(i).Top = Galaga(x).Top + Galaga(x).Height
                    eBullet(i).Visible = True
                    i = 20
                End If
            Next i
        End If
    End If
Next x

'Move Enemy Bullets
For x = 0 To 20
    If eBullet(x).Visible = True Then
        eBullet(x).Top = eBullet(x).Top + 100
            If eBullet(x).Top <= (Ship.Top + Ship.Height) And (eBullet(x).Top + eBullet(x).Height) >= (Ship.Top) And eBullet(x).Left <= (Ship.Left + Ship.Width) And (eBullet(x).Left + eBullet(x).Width) >= Ship.Left Then
                eBullet(x).Visible = False
                Call Rico(Ship.Top - 20, Ship.Left + Ship.Width / 2)
                Ship_Shield = Ship_Shield - 1
                For i = 0 To 9
                    If Shield(i).Visible = True Then
                        Shield(i).Visible = False
                        i = 9
                    End If
                Next i
                If Ship_Shield <= 0 Then
                    Call Explo(Ship.Top, Ship.Left)
                    Call New_Ship
                End If
            End If
        If eBullet(x).Top > 3135 Then eBullet(x).Visible = False
    End If
Next x

For x = 40 To 49
    If Galaga(x).Visible = True Then
        'Bonus During BOSS
        Dim temp As Integer
        temp = Rnd * Bonus_Chance
        If temp = 1 Then
            Call Make_Bonus((Rnd * 2300) + 900, -100)
        End If
    End If
Next x

End Sub

Private Sub lblMyname_Click()   'resume game
TimeStop
TimeStart
End Sub

Private Sub mnAbout2_Click()    'About me
    Call TimeStop
    lblMyname.Visible = True
End Sub

Private Sub mnHow_Click()       '(F1) How to PLay
    Call TimeStop
    rtb1.Left = 0
    rtb1.Top = 0
    rtb1.Visible = True
End Sub

Private Sub mnPause_Click()     '(F3) Paused
    Call TimeStop
    ANS = MsgBox("Click OK to continue", vbOKOnly, "Pause")
    Call TimeStart
End Sub

Private Sub mnQuit_Click()      '(F4) Quits
    End
End Sub

Private Sub Form_Load()
    
    With BlackWinter
        .Height = 3780   'resize form
        .Width = 3375
    End With
    Call HideAll
    Call TimeStop
End Sub

Private Sub mnNew_Click()       '(F2) New game
lblMyname.Visible = False
lblWeaponLaser(0).Visible = True
    Bullet_Speed = 60
    Bullet_Damage = 100
    Bullet_Type = 0
    
    Gun_Lock = 0
    GunLock_L.Visible = True
    GunLock_R.Visible = True
    
    Ship_Speed = 80
    Ship_Dir_N = False  'default stationary
    Ship_Dir_S = False
    Ship_Dir_E = False
    Ship_Dir_W = False
    
    Ship_Ammo = 10
    Ship_Shield = 10
    Ship_Fuel = 10
    
    Ship_Stats(0) = 0   'level
    Ship_Stats(1) = 0   'score
    Ship_Stats(2) = 3   'lives
    Ship_Stats(3) = 3   'nukes

    ShowStats.Interval = 100
    BulletMove.Interval = 20
    ShipMove.Interval = 40
    FuelMove.Interval = 15000
    DistMove.Interval = 5
    GalagaMove.Interval = 40
    ReloadMove.Interval = 500
    StatusMove.Interval = 3000
    EBull.Interval = 100
        
    Ship.Left = 1800    'place ship at default place
    Ship.Top = 2520
    
    For i = 0 To 49 'hide from screen
        Galaga(i).Top = -360
        Galaga(i).Left = -360
    Next i
    
    For i = 0 To 9  'hide from screen
        Bullet(i).Top = -360
        Bullet(i).Left = -360
    Next i
    
    n = NukeStop - 1    'default nuke doesnt run
    Protection.Interval = 1
    Bonus_Chance = 25
    
    For x = 0 To 3
        lblStatus(x) = ""
    Next x

    Send_Msg ("> Ship Launched ...")
    Send_Msg ("> Good Luck Pilot!")
    
    Call ShowAll
    Call TimeStop
    Call New_Level
End Sub

Private Sub New_Level()
    BossAppear = False
    Boss_BARback.Visible = False
    Boss_BARfore.Visible = False
    Boss_BARfore.Left = 1540
    Boss_BARfore.Width = 810
    
    Ship_Stats(0) = Ship_Stats(0) + 1   'LEvel +1
    
    Indic.Left = 120
    
    TimeTemp(5) = NukeTimer.Interval
    Call TimeStart
    DistMove.Interval = 500
End Sub

Private Sub Form_KeyDown(Key As Integer, Shift As Integer)

    'ship Movements.....................................

If Key = 40 Then Ship_Dir_S = True
If Key = 37 Then Ship_Dir_W = True
If Key = 39 Then Ship_Dir_E = True
If Key = 38 Then Ship_Dir_N = True

    'ship WEAPONS.......................................
If Key = 90 Then
    If Bullet_Type = 0 Or Bullet_Type = 1 Then
        GunLock_L.Visible = True
        GunLock_R.Visible = True
        Call Shoot
    Else:
        If Gun_Lock = 0 Then
            Gun_Lock = 1
            GunLock_L.Visible = False
            GunLock_R.Visible = True
            Call Shoot
        End If
    End If
End If

If Key = 88 Then
    If Bullet_Type = 0 Or Bullet_Type = 1 Then
        GunLock_L.Visible = True
        GunLock_R.Visible = True
        Call Shoot
    Else:
        If Gun_Lock = 1 Then
            Gun_Lock = 0
            GunLock_L.Visible = True
            GunLock_R.Visible = False
            Call Shoot
        End If
    End If
End If

If Key = 86 Then        'Nuke Bullets
    Call Nuke
End If

If Key = 67 Then        'Use fuel to reload bullets
    If Ship_Fuel = 0 Then Call FuelMove_Timer
        Call Send_Msg("> Ammo Reloaded ...")
        Ship_Fuel = Ship_Fuel - 1
        Ship_Ammo = 10
        
        For i = 0 To 9
            Ammo(i).Visible = True
        Next i
    
        For i = 0 To 9
            If Fuel(i).Visible = True Then
                Fuel(i).Visible = False
                i = 9
            End If
        Next i
End If
End Sub
Private Sub Form_KeyUp(Key As Integer, Shift As Integer)
If Key = 40 Then Ship_Dir_S = False 'stops ship movement
If Key = 37 Then Ship_Dir_W = False
If Key = 39 Then Ship_Dir_E = False
If Key = 38 Then Ship_Dir_N = False
End Sub

Private Sub Shoot()
If Ship_Ammo > 0 Then
    
    Ship_Ammo = Ship_Ammo - 1       'Reduce stock
    For x = 0 To 9
        If Bullet(x).Visible = False Then  'Draws Bullet
            Bullet(x).Visible = True
            Bullet(x).Left = ((Ship.Left + 180) - (Bullet(x).Width / 2))
            Bullet(x).Top = ((Ship.Top + 100) - (Bullet(x).Height / 2))
                'VULCAN Bullets
            If Bullet_Type = 1 Then
                If x Mod 2 = 0 Then Bullet(x).Left = Bullet(x).Left + 50
                If x Mod 2 = 1 Then Bullet(x).Left = Bullet(x).Left - 50
            End If
            x = 9
        End If
           
    Next x
        For i = 0 To 9
        If Ammo(i).Visible = True Then     'Removes ammo Bar
            Ammo(i).Visible = False
            i = 9
        End If
        Next i
End If
End Sub

Private Sub Nuke()      'SUPER BULLETS
If Ship_Stats(3) > 0 Then
    
    Call Send_Msg("> Let's Nuke em...")

    
    lblNukeFore.Height = 735
    lblNukeFore.Top = 2160
    
    Ship_Stats(3) = Ship_Stats(3) - 1

    For i = 0 To 9
        BulletTemp(0) = Bullet(i).Width
        BulletTemp(1) = Bullet(i).Height
        BulletTemp(2) = Bullet_Speed
        BulletTemp(3) = Bullet_Damage
            
    'SUPREME LASER
        If Bullet_Type = 0 Then
            Bullet(i).Width = 30
            Bullet(i).Height = 300
            Bullet(i).BackColor = &HFFFFC0
            Bullet_Speed = 180
            Bullet_Damage = 400
            If i = 9 Then Call Send_Msg("> SUPREME LASER")
        End If
    
    'GATTLING CANNON
        If Bullet_Type = 1 Then
            Bullet(i).Width = 30
            Bullet(i).Height = 80
            Bullet(i).BackColor = &H8080FF
            Bullet(i).BorderStyle = 1
            Bullet_Speed = 300
            Bullet_Damage = 300
            If i = 9 Then Call Send_Msg("> GATTLING CANNON")
        End If
    
    'NUCLEAR BOMB
        If Bullet_Type = 2 Then
            Bullet(i).Width = 200
            Bullet(i).Height = 280
            Bullet(i).BackColor = &H80FF&
            Bullet(i).BorderStyle = 1
            Bullet_Speed = 50
            Bullet_Damage = 1000
            If i = 9 Then Call Send_Msg("> NUCLEAR BOMB")
        End If
    
    'SONIC WAVES
        If Bullet_Type = 3 Then
            Bullet(i).Width = 800
            Bullet(i).Height = 70
            Bullet(i).BackColor = &HFF00FF
            Bullet(i).BorderStyle = 1
            Bullet_Speed = 90
            Bullet_Damage = 500
            If i = 9 Then Call Send_Msg("> SONIC WAVES")
        End If
    
    'MAD HURRICANE
        If Bullet_Type = 4 Then
            Bullet(i).Width = 80
            Bullet(i).Height = 400
            Bullet(i).BackColor = &H80FF80
            Bullet(i).BorderStyle = 1
            Bullet_Speed = 130
            Bullet_Damage = 400
            If i = 9 Then Call Send_Msg("> MAD HURRICANE")
        End If
    
    Next i
    
    n = 0   'initialize start nuke
    
    lblNukeBack.Visible = True
    lblNukeFore.Visible = True
    NukeTimer.Interval = 250
    NukeEM.Visible = True
End If
End Sub

Private Sub NukeTimer_Timer()   'Stops nuke
n = n + 1

If n Mod 2 = 0 Then NukeEM.Visible = False 'Nuke em PIC Blinking
If n Mod 2 = 1 Then NukeEM.Visible = True

'Nuke em bar
If lblNukeFore.Height > 19 Then lblNukeFore.Height = lblNukeFore.Height - (735 / NukeStop)
If lblNukeFore.Height > 19 Then lblNukeFore.Top = lblNukeFore.Top + (735 / NukeStop)

If n = NukeStop Then    'Nuke ENDS...
    n = 0
    For i = 0 To 9
        Bullet(i).Width = BulletTemp(0) 'Restore Defaults
        Bullet(i).Height = BulletTemp(1)
        Bullet_Speed = BulletTemp(2)
        Bullet_Damage = BulletTemp(3)
    Next i
        For i = 0 To 9
        Bullet(i).BorderStyle = 0
        Next i
    lblNukeFore.Visible = False
    lblNukeBack.Visible = False
    NukeTimer.Interval = 0
End If
End Sub

Private Sub End_Game()  'Game Over
    Call TimeStop
    Call HideAll
    ANS = MsgBox("Mission Failed. No more ships left.", vbOKOnly, "GAME OVER")
    lblMyname.Visible = True
End Sub

Private Sub TimeStart() 'Resumes Time
    ShowStats.Enabled = True
    BulletMove.Enabled = True
    ShipMove.Enabled = True
    FuelMove.Enabled = True
    DistMove.Enabled = True
    EnemyMove.Enabled = True
    NukeTimer.Enabled = True
    GalagaMove.Enabled = True
    ExploMove.Enabled = True
    RicoMove.Enabled = True
    Protection.Enabled = True
    ReloadMove.Enabled = True
    EBull.Enabled = True
End Sub

Private Sub TimeStop()  'Stops Time
    ShowStats.Enabled = False
    BulletMove.Enabled = False
    ShipMove.Enabled = False
    FuelMove.Enabled = False
    DistMove.Enabled = False
    EnemyMove.Enabled = False
    NukeTimer.Enabled = False
    GalagaMove.Enabled = False
    ExploMove.Enabled = False
    RicoMove.Enabled = False
    Protection.Enabled = False
    ReloadMove.Enabled = False
    EBull.Enabled = False
End Sub

Private Sub ShowAll()       'Show all Stats
    For x = 0 To (Ship_Ammo - 1)
        Ammo(x).Visible = True
    Next x
    
    For x = 0 To (Ship_Shield - 1)
        Shield(x).Visible = True
    Next
    
    For x = 0 To (Ship_Fuel - 1)
        Fuel(x).Visible = True
    Next

Indic.Visible = True
Ship.Visible = True
lblGun_Lock.Visible = True

End Sub

Private Sub HideAll()

'hide All small items
For x = 0 To 9
    Bullet(x).Visible = False
    Ammo(x).Visible = False
    Shield(x).Visible = False
    Fuel(x).Visible = False
Next

Prot_Circle.Visible = False
Indic.Visible = False
Ship.Visible = False
lblGun_Lock.Visible = False
GunLock_L.Visible = False
GunLock_R.Visible = False

For x = 0 To 4  'hides weapons label
    lblWeaponLaser(x).Visible = False
    PWup(x).Visible = False
Next x

For x = 0 To 49 'hides enemy
    Galaga(x).Visible = False
Next x

For x = 0 To 20
    eBullet(x).Visible = False
Next

End Sub

Private Sub ReloadMove_Timer()  'this timer reloads bullets
            For i = 9 To 0 Step -1
                If Ammo(i).Visible = False Then
                    Ammo(i).Visible = True         'reload bar
                    Ship_Ammo = Ship_Ammo + 1      'reload stock
                    i = 0
                End If
            Next i
End Sub

Private Sub RTB1_Click()    'resumes game
rtb1.Visible = False
Call TimeStart
End Sub

Private Sub ShipMove_Timer()    'Ship Timer

    If Ship_Protect = True Then
        Prot_Circle.Top = Ship.Top
        Prot_Circle.Left = Ship.Left
        Prot_Circle.Visible = True
    End If

    'MOVES ship in direction set
    If Ship_Dir_S = True Then
        Ship.Top = Ship.Top + Ship_Speed
    End If
    If Ship_Dir_W = True Then
        Ship.Left = Ship.Left - Ship_Speed
    End If
    If Ship_Dir_E = True Then
        Ship.Left = Ship.Left + Ship_Speed
    End If
    If Ship_Dir_N = True Then
        Ship.Top = Ship.Top - Ship_Speed
    End If
    
    'STOPS ship is crashes into wall
    If Ship.Left <= 840 Then
        Ship.Left = 850
        Ship_Dir_W = False
    End If
    If Ship.Left >= 2880 Then
        Ship.Left = 2870
        Ship_Dir_E = False
    End If
    If Ship.Top >= 2760 Then
        Ship.Top = 2750
        Ship_Dir_S = False
    End If
    If Ship.Top <= 120 Then
        Ship.Top = 130
        Ship_Dir_N = False
    End If


    'if galaga hits ship
        For b = 0 To 49
            If Galaga(b).Visible Then
                If Ship_Protect = False And Ship.Top <= (Galaga(b).Top + Galaga(b).Height) And (Ship.Top + Ship.Height) >= (Galaga(b).Top) And Ship.Left <= (Galaga(b).Left + Galaga(b).Width) And (Ship.Left + Ship.Width) >= Galaga(b).Left Then
                    Galaga_Life(b) = Galaga_Life(b) - 500   'Inflicts 500 to enemy
                    
                        Ship_Shield = Ship_Shield - 3   'Lose 3 shields for a CRASH
                        For x = 1 To 3
                            For i = 0 To 9
                                If Shield(i).Visible = True Then
                                    Shield(i).Visible = False
                                    i = 9
                                End If
                            Next i
                        Next x
                        
                        If Ship_Shield <= 0 Then
                            Call Explo(Ship.Top, Ship.Left)
                            Call New_Ship
                        End If
                    
                    If Galaga_Life(b) <= 0 Then
                        Galaga(b).Visible = False
                        Ship_Stats(1) = Ship_Stats(1) + Galaga_Worth(b)
                        Call Explo(Galaga(b).Top, Galaga(b).Left)
                    End If
                
                End If
            End If
        Next b

If Ship_Stats(2) <= 0 Then Call End_Game    'NO MORE LIVES

End Sub

Private Sub New_Ship()  'A BRAND NEW SHIP
    
    Call Send_Msg("> Ship Destroyed ...")
    
    Ship_Stats(2) = Ship_Stats(2) - 1   'Lives - 1
                            
    For i = 0 To 9
        Ammo(i).Visible = True
        Shield(i).Visible = True
        Fuel(i).Visible = True
    Next i
    
    Ship_Ammo = 10
    Ship_Fuel = 10
    Ship_Shield = 10
    Ship.Left = 1800
    Ship.Top = 2520
                    
    Protection.Interval = 1
    
    Ship_Dir_N = False
    Ship_Dir_S = False
    Ship_Dir_E = False
    Ship_Dir_W = False
End Sub

Private Sub Protection_Timer()
    If Protection.Interval = 1 Then
        Ship_Protect = True
        Protection.Interval = 4000
    ElseIf Protection.Interval = 4000 Then
        Ship_Protect = False
        Prot_Circle.Visible = False
        Protection.Interval = 0
        Send_Msg ("> Blue Shield Faded...")
    End If
End Sub

Private Sub ShowStats_Timer()       'Update All Stats
For x = 0 To 3
    Stats(x).Caption = Str(Ship_Stats(x))
Next x

For x = 0 To 4          'Bullet type label
    If Bullet_Type = x Then
        lblWeaponLaser(x).Visible = True
    Else: lblWeaponLaser(x).Visible = False
    End If
Next x

If NukeTimer.Interval = 0 Then
'-----------------BULLET TYPES --------------------
'LASER
    If Bullet_Type = 0 Then
        For i = 0 To 9
            Bullet(i).Width = 15
            Bullet(i).Height = 150
            Bullet(i).BackColor = &HFFFFC0
            Bullet_Speed = 110
            Bullet_Damage = 100
            ReloadMove.Interval = 300
        Next i
    End If
'VULCAN
    If Bullet_Type = 1 Then
        For i = 0 To 9
            Bullet(i).Width = 30
            Bullet(i).Height = 80
            Bullet(i).BackColor = &H8080FF
            Bullet_Speed = 250
            Bullet_Damage = 55
            ReloadMove.Interval = 130
        Next i
    End If
'NAPALM
    If Bullet_Type = 2 Then
        For i = 0 To 9
            Bullet(i).Width = 100
            Bullet(i).Height = 140
            Bullet(i).BackColor = &H80FF&
            Bullet_Speed = 40
            Bullet_Damage = 370
            ReloadMove.Interval = 600
            Bullet(i).BorderStyle = 1
        Next i
        Else
        For i = 0 To 9
        Bullet(i).BorderStyle = 0
        Next i
    End If
'WAVE
    If Bullet_Type = 3 Then
        For i = 0 To 9
            Bullet(i).Width = 500
            Bullet(i).Height = 40
            Bullet(i).BackColor = &HFF00FF
            Bullet_Speed = 70
            Bullet_Damage = 190
            ReloadMove.Interval = 350
        Next i
    End If
'TYPHOON
    If Bullet_Type = 4 Then
        For i = 0 To 9
            Bullet(i).Width = 40
            Bullet(i).Height = 300
            Bullet(i).BackColor = &H80FF80
            Bullet_Speed = 100
            Bullet_Damage = 160
            ReloadMove.Interval = 300
        Next i
    End If

End If
End Sub

Private Sub BulletMove_Timer()  'Bullet Timer

If BossAppear = True Then
    For x = 40 To 49
        If Galaga(x).Visible = True Then
            Boss_BARfore.Width = (Galaga_Life(x) / Galaga_Boss_Bar) * Boss_BARback.Width
        End If
    Next x
End If

For x = 0 To 9
    If Bullet(x).Visible Then
        Bullet(x).Top = Bullet(x).Top - Bullet_Speed    'moves bullet
        
      
                'for TYPHOON Bullets
                If Bullet_Type = 4 Then Bullet(x).Left = Bullet(x).Left + ((Rnd * 160) - 80)
        
        If Bullet(x).Top <= -120 Then              'bullet out of screen
            Bullet(x).Visible = False              'removes bullet
        End If
        
    'if bullet hits GALAGA
        For b = 0 To 49
            If Galaga(b).Visible = True Then
                If Bullet(x).Top <= (Galaga(b).Top + Galaga(b).Height) And (Bullet(x).Top + Bullet(x).Height) >= (Galaga(b).Top) And Bullet(x).Left <= (Galaga(b).Left + Galaga(b).Width) And (Bullet(x).Left + Bullet(x).Width) >= Galaga(b).Left Then
                    
                    Bullet(x).Visible = False
                    Call Rico((Galaga(b).Top + Galaga(b).Height - 50), (Galaga(b).Left + Galaga(b).Width / 3))
                    
                    Galaga_Life(b) = Galaga_Life(b) - Bullet_Damage
                    If Galaga_Life(b) <= 0 Then
                        Galaga(b).Visible = False
                        Ship_Stats(1) = Ship_Stats(1) + Galaga_Worth(b)
                        
                        If b < 40 Then
                            Call Make_Bonus(Galaga(b).Left, Galaga(b).Top)
                        End If
                        
                        Call Explo(Galaga(b).Top, Galaga(b).Left)
                            If b >= 40 Then 'killed boss!!!
                                If b = 49 Then
                                    TimeStop
                                    ANS = MsgBox("You've COMPLETED Black Winter !!", vbOKOnly, "CONGRATULATIONS!!!")
                                    ANS = MsgBox("THE END", vbOKOnly, "CONGRATULATIONS!!!")
                                    ANS = MsgBox("Black Winter II Coming Soon ...", vbOKOnly, "CONGRATULATIONS!!!")
                                    Call End_Game
                                Else:
                                    Send_Msg ("> Target Eliminated")
                                    Call New_Level
                                End If
                            End If
                    End If
                End If
            End If
        Next b
    End If
Next x

For x = 0 To 4      'Bonus MOVEMENTS
    If PWup(x).Visible = True Then
        PWup(x).Top = PWup(x).Top + 15
        If PWup(x).Top > 3135 Then PWup(x).Visible = False
            If PWup(x).Top <= (Ship.Top + Ship.Height) And (PWup(x).Top + PWup(x).Height) >= (Ship.Top) And PWup(x).Left <= (Ship.Left + Ship.Width) And (PWup(x).Left + PWup(x).Width) >= Ship.Left Then
                PWup(x).Visible = False
                
                If x = 0 Then
                    Ship_Ammo = 10
                    For i = 9 To 0 Step -1
                        If Ammo(i).Visible = False Then
                            Ammo(i).Visible = True         'reload bar
                        End If
                    Next i
                i = (Rnd * 4)
                If i = Bullet_Type Then i = i + 1
                If i = 5 Then i = 1
                Bullet_Type = i
                Send_Msg ("> Ammo Pickup")
                    If Bullet_Type = 0 Or Bullet_Type = 1 Then
                        Send_Msg ("> Dual Shot Engaged")
                    Else
                        Send_Msg ("> Mono Shot Engaged")
                    End If
                                
                ElseIf x = 1 Then
                    Ship_Shield = Ship_Shield + 2
                    If Ship_Shield > 10 Then Ship_Shield = 10
                    For b = 0 To 1
                        For i = 9 To 0 Step -1
                            If Shield(i).Visible = False Then
                                Shield(i).Visible = True         'reload bar
                                i = 0
                            End If
                        Next i
                    Next b
                    Send_Msg ("> Shield Pickup")

                ElseIf x = 2 Then
                    Ship_Fuel = Ship_Fuel + 3
                    If Ship_Fuel > 10 Then Ship_Fuel = 10
                    For b = 0 To 2
                        For i = 9 To 0 Step -1
                            If Fuel(i).Visible = False Then
                                Fuel(i).Visible = True         'reload bar
                                i = 0
                            End If
                        Next i
                    Next b
                    Send_Msg ("> Fuel Pickup")

                ElseIf x = 3 Then
                    Ship_Stats(2) = Ship_Stats(2) + 1
                    Send_Msg ("> Extra Ship")

                ElseIf x = 4 Then
                    Ship_Stats(3) = Ship_Stats(3) + 1
                    Send_Msg ("> Nuke Pickup")
                End If

            End If
    End If
Next x

End Sub
Private Sub Make_Bonus(Left, Top)
                            Dim TempBonus As Integer
                            TempBonus = (Rnd * 50)
                            
                            'Bonus Life
                            If TempBonus = 1 And PWup(3).Visible = False Then
                                With PWup(3)
                                .Visible = True
                                .Left = Left
                                .Top = Top
                                End With
                            End If
                            'Bonus Weapon
                            If TempBonus > 14 And TempBonus < 18 And PWup(0).Visible = False Then
                                With PWup(0)
                                .Visible = True
                                .Left = Left
                                .Top = Top
                                End With
                            End If
                            'Bonus Nuke
                            If TempBonus > 1 And TempBonus < 3 And PWup(4).Visible = False Then
                                With PWup(4)
                                .Visible = True
                                .Left = Left
                                .Top = Top
                                End With
                            End If
                            'Bonus Shield
                            If TempBonus > 2 And TempBonus < 6 And PWup(1).Visible = False Then
                                With PWup(1)
                                .Visible = True
                                .Left = Left
                                .Top = Top
                                End With
                            End If
                            'Bonus Fuel
                            If TempBonus > 5 And TempBonus < 15 And PWup(2).Visible = False Then
                                With PWup(2)
                                .Visible = True
                                .Left = Left
                                .Top = Top
                                End With
                            End If

End Sub

Private Sub Rico(Rico_x, Rico_y)     'Bullets Ricochet
    picRico.Top = Rico_x
    picRico.Left = Rico_y
    picRico.Visible = True
    RicoMove.Interval = 19
    

    
End Sub
Private Sub RicoMove_Timer()    'BULLET RICOCHET
    picRico.Visible = False
    RicoMove.Interval = 0
End Sub
Private Sub Explo(Explo_x, Explo_y) 'Explosion animation
    e = 0
    PicExplo1.Top = Explo_x 'set explosion sprite to explode site
    PicExplo1.Left = Explo_y
    picExplo2.Top = Explo_x
    picExplo2.Left = Explo_y
    picExplo3.Top = Explo_x
    picExplo3.Left = Explo_y
    picExplo4.Top = Explo_x
    picExplo4.Left = Explo_y
    ExploMove.Interval = 50
    
End Sub

Private Sub ExploMove_Timer()
    If e = 0 Then
        PicExplo1.Visible = True
        e = e + 1
                
       
        ElseIf e = 1 Then       'slowly draws out explosion animation
            PicExplo1.Visible = False
            picExplo2.Visible = True
            e = e + 1
            ElseIf e = 2 Then
            picExplo2.Visible = False
                picExplo3.Visible = True
                e = e + 1
                ElseIf e = 3 Then
                    picExplo3.Visible = False
                    picExplo4.Visible = True
                    e = e + 1
                    ElseIf e = 4 Then
                        picExplo4.Visible = False
                        ExploMove.Interval = 0
    End If
End Sub

Private Sub FuelMove_Timer()        'Fuel Timer
    Ship_Fuel = Ship_Fuel - 1
    For x = 0 To 9
        If Fuel(x).Visible = True Then
            Fuel(x).Visible = False
            x = 9
        End If
    Next x

    If Ship_Fuel <= 0 Then      'Empty Fuel... lose 1 ship
        Call TimeStop
        ANS = MsgBox("Your fuel gauge is empty", vbOKOnly, "Ship Lost")
        Call New_Ship
        Call TimeStart
    End If
End Sub
Private Sub Summon_Boss()
    BossAppear = True
    Boss_BARfore.Visible = True
    Boss_BARback.Visible = True
    Call Send_Msg("> Incoming Large Ship!")
    
    Select Case (Ship_Stats(0))
        Case Is = 1:
            Call Galaga_Wave(9999, 8, 47, 47, 1440, -760, 30, 12000, 1000, 100)
            Send_Msg ("> BOSS : Fallen Mask")
            Send_Msg ("> BOSS : Fallen Mask")
        Case Is = 2:
            Call Galaga_Wave(9999, 8, 43, 43, 1440, -760, 40, 20000, 2000, 150)
            Send_Msg ("> BOSS : Stray Capt")
            Send_Msg ("> BOSS : Stray Capt")
        Case Is = 3:
            Call Galaga_Wave(9999, 8, 48, 48, 1440, -760, 80, 17000, 2000, 170)
            Send_Msg ("> BOSS : Marsh-Yellow")
            Send_Msg ("> BOSS : Marsh-Yellow")
            Bonus_Chance = 20
        Case Is = 4:
            Call Galaga_Wave(9999, 8, 40, 40, 1440, -760, 20, 50000, 3000, 180)
            Send_Msg ("> BOSS : Hydro Knight")
            Send_Msg ("> BOSS : Hydro Knight")
        Case Is = 5:
            Call Galaga_Wave(9999, 8, 42, 42, 1440, -760, 150, 30000, 4000, 250)
            Send_Msg ("> BOSS : Feral Demon")
            Send_Msg ("> BOSS : Feral Demon")
            Bonus_Chance = 15
        Case Is = 6:
            Call Galaga_Wave(9999, 8, 45, 45, 1440, -760, 80, 40000, 6000, 190)
            Send_Msg ("> BOSS : Frozen Wish")
            Send_Msg ("> BOSS : Frozen Wish")
            Bonus_Chance = 10
        Case Is = 7:
            Call Galaga_Wave(9999, 8, 44, 44, 1440, -760, 90, 50000, 6000, 200)
            Send_Msg ("> BOSS : Bloody Sunday")
            Send_Msg ("> BOSS : Bloody Sunday")
            Bonus_Chance = 10
        Case Is = 8:
            Call Galaga_Wave(9999, 8, 41, 41, 1440, -760, 100, 70000, 8000, 240)
            Send_Msg ("> BOSS : One Eye Lord")
            Send_Msg ("> BOSS : One Eye Lord")
            Bonus_Chance = 5
        Case Is = 9:
            Call Galaga_Wave(9999, 8, 46, 46, 1440, -760, 100, 100000, 8000, 270)
            Send_Msg ("> BOSS : Space Visitor")
            Send_Msg ("> BOSS : Space Visitor")
        Case Is = 10:
            Call Galaga_Wave(9999, 8, 49, 49, 1440, -760, 150, 1000000, 100000, 500)
            Send_Msg ("> BOSS : Unholy Shade")
            Send_Msg ("> BOSS : Unholy Shade")
    End Select
    
End Sub


Private Sub DistMove_Timer()    'Distance Indicator
    If Indic.Visible = True Then
    Indic.Left = Indic.Left + 2
        If Indic.Left >= 600 Then
            Call Summon_Boss
            DistMove.Interval = 0
        End If
    End If
    
'---------------------Level Design--------------------
    'Move Pats available    :   2426 2624 2222 3322 1122 3332
    '                           1112 3212 1232 4444 6666 9999 (boss)
    
    'Directions available   :   4 6 7 8 9
    'Start Pos x            :   840 - 2880
    'Start Pos y            :     0 - 2760  (default : -360)
    'Start - End Num        :     0 -   39  (40-49 bosses)
    'MAX SPEED              :    30 -  100
    'Health                 :   100 -  Max
    'Damage                 :     0 -  99    (shoot frequency)
        
    '*** indic (0-720) summon event must be EVEN number ***
            
    If Ship_Stats(0) = 1 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(2624, 8, 0, 2, 1200, -360, 30, 100, 10, 10)
        Case Is = 40: Call Galaga_Wave(2222, 8, 6, 7, 3000, -360, 30, 100, 10, 10)
        Case Is = 60: Call Galaga_Wave(2222, 9, 3, 5, 1500, -360, 30, 100, 10, 10)
        Case Is = 80: Call Galaga_Wave(1112, 8, 8, 9, 3000, -360, 40, 100, 10, 10)
        Case Is = 100: Call Galaga_Wave(2222, 8, 9, 10, 1000, -360, 40, 100, 10, 10)
        Case Is = 140: Call Galaga_Wave(2222, 8, 0, 3, 1000, -360, 30, 100, 10, 10)
        Case Is = 160: Call Galaga_Wave(2222, 8, 4, 7, 2600, -360, 30, 100, 10, 10)
        Case Is = 180: Call Galaga_Wave(2222, 8, 8, 11, 1000, -360, 25, 100, 10, 10)
        Case Is = 210: Call Galaga_Wave(2222, 8, 16, 19, 1500, -360, 25, 100, 10, 10)
        Case Is = 250: Call Galaga_Wave(2222, 8, 12, 15, 2000, -360, 30, 100, 10, 10)
        Case Is = 260: Call Galaga_Wave(2222, 8, 0, 2, 1100, -360, 30, 100, 10, 10)
        Case Is = 270: Call Galaga_Wave(2222, 8, 3, 5, 2000, -360, 30, 100, 10, 10)
        Case Is = 310: Call Galaga_Wave(2222, 8, 12, 15, 2900, -360, 30, 100, 10, 10)
        Case Is = 330: Call Galaga_Wave(2222, 8, 6, 8, 2965, -360, 40, 150, 15, 10)
        Case Is = 350: Call Galaga_Wave(2222, 8, 8, 11, 1040, -360, 40, 150, 15, 10)
        Case Is = 370: Call Galaga_Wave(2222, 8, 0, 3, 1000, -360, 30, 150, 15, 10)
        Case Is = 390: Call Galaga_Wave(2222, 8, 8, 12, 2000, -360, 30, 150, 15, 10)
        Case Is = 410: Call Galaga_Wave(2222, 8, 4, 7, 2800, -360, 30, 150, 15, 10)
        Case Is = 440: Call Galaga_Wave(2624, 6, 16, 19, 1200, -360, 30, 200, 15, 10)
        Case Is = 450: Call Galaga_Wave(2222, 8, 0, 3, 1000, -360, 30, 100, 10, 10)
        End Select
    End If
    

    If Ship_Stats(0) = 2 Then
        Select Case (Indic.Left - 120)
        Case Is = 30: Call Galaga_Wave(2624, 8, 0, 2, 1200, -360, 30, 100, 10, 10)
        Case Is = 50: Call Galaga_Wave(3322, 7, 3, 5, 1400, -360, 20, 200, 15, 15)
        Case Is = 70: Call Galaga_Wave(2426, 4, 6, 8, 2600, -360, 40, 200, 15, 15)
        Case Is = 90: Call Galaga_Wave(3332, 7, 17, 17, 1300, -360, 30, 250, 20, 10)
        Case Is = 110: Call Galaga_Wave(1112, 7, 18, 18, 3000, -360, 30, 250, 20, 10)
        Case Is = 130: Call Galaga_Wave(2222, 6, 0, 4, 1400, -360, 20, 250, 15, 10)
        Case Is = 170: Call Galaga_Wave(3212, 7, 9, 11, 1400, -360, 30, 200, 15, 10)
        Case Is = 190: Call Galaga_Wave(1222, 9, 12, 15, 2000, -360, 30, 200, 15, 10)
        Case Is = 210: Call Galaga_Wave(2222, 6, 5, 8, 1400, -360, 20, 300, 15, 10)
        Case Is = 230: Call Galaga_Wave(6666, 9, 0, 4, 100, 1800, 30, 150, 15, 10)
        Case Is = 250: Call Galaga_Wave(6666, 9, 8, 11, 100, 1800, 30, 200, 20, 15)
        Case Is = 270: Call Galaga_Wave(3332, 7, 17, 17, 1500, -360, 30, 250, 20, 10)
        Case Is = 290: Call Galaga_Wave(2624, 8, 17, 17, 1700, -360, 40, 400, 30, 15)
        Case Is = 310: Call Galaga_Wave(2624, 8, 18, 18, 1900, -360, 40, 400, 30, 15)
        Case Is = 330: Call Galaga_Wave(2222, 8, 6, 8, 2365, -360, 40, 250, 25, 15)
        Case Is = 350: Call Galaga_Wave(2222, 8, 8, 11, 1040, -360, 40, 250, 25, 15)
        Case Is = 370: Call Galaga_Wave(2222, 8, 0, 3, 1700, -360, 30, 250, 25, 15)
        Case Is = 390: Call Galaga_Wave(2222, 6, 0, 4, 1400, -360, 20, 250, 15, 10)
        Case Is = 410: Call Galaga_Wave(2624, 8, 0, 2, 1200, -360, 30, 100, 10, 10)
        Case Is = 430: Call Galaga_Wave(3322, 7, 3, 5, 1400, -360, 20, 200, 15, 15)
        End Select
    End If
    
    If Ship_Stats(0) = 3 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(3332, 7, 0, 3, 1200, -360, 40, 250, 25, 15)
        Case Is = 40: Call Galaga_Wave(1112, 9, 4, 7, 2700, -360, 40, 250, 25, 15)
        Case Is = 60: Call Galaga_Wave(2222, 6, 8, 11, 1300, -360, 30, 300, 25, 15)
        Case Is = 80: Call Galaga_Wave(2624, 6, 12, 15, 2800, -360, 40, 300, 25, 15)
        Case Is = 100: Call Galaga_Wave(2426, 4, 16, 19, 2200, -360, 40, 300, 25, 15)
        Case Is = 120: Call Galaga_Wave(2222, 8, 20, 23, 2000, -360, 30, 300, 25, 15)
        Case Is = 130: Call Galaga_Wave(9999, 8, 24, 24, 2000, -360, 40, 800, 60, 50)
        Case Is = 140: Call Galaga_Wave(1112, 9, 4, 7, 2600, -360, 40, 250, 25, 15)
        Case Is = 160: Call Galaga_Wave(3212, 9, 8, 11, 2000, -360, 30, 400, 35, 15)
        Case Is = 180: Call Galaga_Wave(1232, 7, 12, 15, 1400, -360, 30, 400, 35, 15)
        Case Is = 200: Call Galaga_Wave(4444, 8, 0, 5, 3400, 2000, 30, 250, 25, 15)
        Case Is = 220: Call Galaga_Wave(6666, 8, 8, 13, 800, 2000, 30, 250, 25, 15)
        Case Is = 250: Call Galaga_Wave(3332, 7, 0, 3, 1200, -360, 40, 450, 25, 15)
        Case Is = 250: Call Galaga_Wave(1112, 9, 4, 7, 2700, -360, 40, 450, 35, 15)
        Case Is = 280: Call Galaga_Wave(1112, 9, 8, 11, 2700, -360, 40, 500, 35, 15)
        Case Is = 300: Call Galaga_Wave(3212, 9, 12, 15, 2600, -360, 30, 500, 35, 15)
        Case Is = 330: Call Galaga_Wave(2222, 8, 0, 3, 1855, -360, 40, 450, 35, 15)
        Case Is = 350: Call Galaga_Wave(2222, 6, 8, 11, 1300, -360, 40, 500, 45, 15)
        Case Is = 370: Call Galaga_Wave(2222, 6, 12, 15, 1300, -360, 40, 500, 45, 15)
        Case Is = 390: Call Galaga_Wave(2222, 6, 16, 19, 1300, -360, 40, 600, 45, 25)
        Case Is = 420: Call Galaga_Wave(6666, 4, 20, 23, 1040, -360, 40, 250, 25, 15)
        Case Is = 440: Call Galaga_Wave(4444, 6, 24, 27, 1000, -360, 30, 250, 25, 15)
        End Select
    End If
    

    If Ship_Stats(0) = 4 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(3332, 7, 0, 3, 1200, -360, 40, 250, 25, 15)
        Case Is = 40: Call Galaga_Wave(1112, 9, 4, 7, 2700, -360, 40, 250, 25, 15)
        Case Is = 60: Call Galaga_Wave(2222, 6, 8, 11, 1300, -360, 30, 300, 25, 15)
        Case Is = 80: Call Galaga_Wave(2624, 6, 12, 15, 2800, -360, 40, 300, 25, 15)
        Case Is = 100: Call Galaga_Wave(2426, 4, 16, 19, 1200, -360, 40, 300, 25, 15)
        Case Is = 120: Call Galaga_Wave(2222, 8, 20, 23, 2000, -360, 30, 300, 25, 15)
        Case Is = 130: Call Galaga_Wave(2222, 8, 3, 3, 1400, -360, 40, 800, 60, 50)
        Case Is = 140: Call Galaga_Wave(1112, 9, 4, 7, 2500, -360, 40, 250, 25, 15)
        Case Is = 160: Call Galaga_Wave(3212, 9, 8, 11, 1800, -360, 30, 400, 35, 15)
        Case Is = 180: Call Galaga_Wave(1232, 7, 12, 15, 1400, -360, 30, 400, 35, 15)
        Case Is = 200: Call Galaga_Wave(4444, 8, 30, 35, 3400, 2000, 30, 250, 25, 15)
        Case Is = 220: Call Galaga_Wave(6666, 8, 8, 13, 800, 2000, 30, 250, 25, 15)
        Case Is = 330: Call Galaga_Wave(2222, 8, 0, 3, 1965, -360, 40, 450, 35, 15)
        Case Is = 350: Call Galaga_Wave(2222, 6, 8, 11, 1800, -360, 40, 500, 45, 15)
        Case Is = 370: Call Galaga_Wave(2222, 6, 12, 15, 1400, -360, 40, 500, 45, 15)
        Case Is = 390: Call Galaga_Wave(2222, 6, 16, 19, 1500, -360, 40, 600, 45, 25)
        Case Is = 420: Call Galaga_Wave(6666, 4, 20, 23, 1040, -360, 40, 250, 25, 15)
        Case Is = 440: Call Galaga_Wave(4444, 6, 24, 27, 3000, -360, 30, 250, 25, 15)
        End Select
    End If
    
    If Ship_Stats(0) = 5 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(2426, 8, 32, 39, 1800, -100, 20, 500, 200, 50)
        Case Is = 60: Call Galaga_Wave(2222, 6, 10, 15, 1200, -200, 20, 100, 105, 35)
        Case Is = 100: Call Galaga_Wave(6666, 6, 20, 24, 100, 500, 25, 200, 125, 55)
        Case Is = 140: Call Galaga_Wave(6666, 6, 25, 29, 100, 800, 25, 200, 135, 55)
        Case Is = 180: Call Galaga_Wave(6666, 6, 30, 34, 100, 1100, 25, 200, 155, 65)
        Case Is = 220: Call Galaga_Wave(2222, 6, 24, 27, 1300, -260, 25, 200, 235, 45)
        Case Is = 260: Call Galaga_Wave(2222, 6, 28, 31, 1400, -260, 25, 200, 235, 45)
        Case Is = 300: Call Galaga_Wave(1112, 8, 35, 38, 2200, -360, 25, 200, 185, 45)
        Case Is = 300: Call Galaga_Wave(3332, 8, 30, 34, 1600, -360, 25, 300, 185, 45)
        Case Is = 340: Call Galaga_Wave(1112, 9, 0, 7, 2600, -60, 25, 300, 435, 65)
        Case Is = 380: Call Galaga_Wave(1112, 9, 8, 13, 2660, -60, 25, 300, 435, 65)
        Case Is = 420: Call Galaga_Wave(1112, 9, 14, 20, 2600, -60, 25, 300, 435, 65)
        Case Is = 440: Call Galaga_Wave(1112, 9, 21, 27, 2600, -60, 25, 300, 43, 65)
        Case Is = 460: Call Galaga_Wave(2426, 4, 28, 29, 2500, -360, 25, 355, 650, 100)
        End Select
    End If

    If Ship_Stats(0) = 6 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(2426, 8, 32, 35, 1800, -100, 35, 800, 500, 50)
        Case Is = 40: Call Galaga_Wave(3332, 7, 0, 3, 1500, -360, 40, 450, 25, 15)
        Case Is = 60: Call Galaga_Wave(9999, 6, 14, 17, 1500, -200, 35, 1800, 605, 35)
        Case Is = 80: Call Galaga_Wave(6666, 6, 28, 30, 800, 500, 35, 800, 625, 55)
        Case Is = 130: Call Galaga_Wave(1112, 9, 4, 7, 2000, -360, 40, 450, 35, 15)
        Case Is = 150: Call Galaga_Wave(1112, 9, 8, 11, 1400, -360, 40, 500, 35, 15)
        Case Is = 170: Call Galaga_Wave(3212, 9, 12, 15, 2500, -360, 30, 500, 35, 15)
        Case Is = 190: Call Galaga_Wave(1112, 8, 36, 37, 1200, -360, 35, 900, 685, 45)
        Case Is = 210: Call Galaga_Wave(3322, 8, 31, 33, 2800, -360, 35, 900, 685, 45)
        Case Is = 230: Call Galaga_Wave(9999, 9, 0, 3, 2400, -60, 35, 800, 2635, 65)
        Case Is = 250: Call Galaga_Wave(1122, 9, 4, 5, 2800, -60, 35, 800, 635, 65)
        Case Is = 270: Call Galaga_Wave(6666, 6, 34, 36, 800, 800, 35, 800, 735, 55)
        Case Is = 290: Call Galaga_Wave(6666, 6, 37, 39, 800, 1100, 35, 800, 755, 65)
        Case Is = 310: Call Galaga_Wave(2222, 6, 24, 27, 1300, -260, 35, 900, 735, 10)
        Case Is = 330: Call Galaga_Wave(2222, 6, 28, 31, 1400, -260, 35, 900, 735, 10)
        Case Is = 360: Call Galaga_Wave(1122, 9, 2, 5, 2800, -60, 30, 800, 735, 10)
        Case Is = 380: Call Galaga_Wave(1122, 9, 6, 9, 2800, -60, 30, 800, 273, 10)
        Case Is = 400: Call Galaga_Wave(2222, 6, 30, 33, 1300, 1000, 5, 2400, 1650, 0)
        Case Is = 420: Call Galaga_Wave(2222, 6, 36, 39, 1530, 1200, 5, 2400, 1650, 0)
        End Select
    End If
    
    If Ship_Stats(0) = 7 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(2624, 8, 0, 3, 1200, -360, 60, 100, 10, 30)
        Case Is = 30:
            Call Galaga_Wave(2222, 8, 3, 7, 3000, -360, 60, 100, 10, 30)
            Send_Msg ("> HIT and RUN!!")
            Send_Msg ("> HIT and RUN!!")
        Case Is = 40: Call Galaga_Wave(1112, 8, 8, 11, 3000, -360, 60, 100, 310, 30)
        Case Is = 50: Call Galaga_Wave(2222, 8, 12, 14, 1000, -360, 60, 100, 310, 30)
        Case Is = 60: Call Galaga_Wave(2222, 8, 15, 18, 1000, -360, 60, 100, 310, 30)
        Case Is = 70: Call Galaga_Wave(2222, 8, 19, 22, 2600, -360, 60, 100, 310, 30)
        Case Is = 80: Call Galaga_Wave(2222, 8, 23, 25, 1000, -360, 75, 100, 310, 30)
        Case Is = 90: Call Galaga_Wave(2222, 8, 25, 30, 1500, -360, 75, 100, 310, 30)
        Case Is = 100: Call Galaga_Wave(2222, 8, 32, 36, 2000, -360, 70, 100, 310, 30)
        Case Is = 130: Call Galaga_Wave(2222, 8, 0, 5, 1100, -360, 70, 100, 310, 30)
        Case Is = 140: Call Galaga_Wave(2222, 8, 6, 9, 2000, -360, 70, 100, 310, 30)
        Case Is = 150: Call Galaga_Wave(2222, 8, 10, 14, 2900, -360, 70, 100, 310, 30)
        Case Is = 170: Call Galaga_Wave(2222, 8, 15, 20, 2965, -360, 70, 150, 315, 30)
        Case Is = 180: Call Galaga_Wave(2222, 8, 21, 24, 1040, -360, 70, 150, 315, 30)
        Case Is = 190: Call Galaga_Wave(2222, 8, 32, 35, 1000, -360, 70, 150, 315, 30)
        Case Is = 200: Call Galaga_Wave(2222, 8, 0, 15, 2000, -360, 70, 150, 315, 30)
        Case Is = 210: Call Galaga_Wave(2222, 8, 16, 19, 2800, -360, 70, 150, 315, 30)
        Case Is = 220: Call Galaga_Wave(2624, 6, 20, 25, 1200, -360, 70, 200, 315, 30)
        Case Is = 230: Call Galaga_Wave(2222, 8, 26, 30, 1000, -360, 70, 100, 330, 30)
        Case Is = 240: Call Galaga_Wave(2222, 8, 32, 25, 2600, -360, 80, 100, 310, 30)
        Case Is = 250: Call Galaga_Wave(2624, 8, 36, 39, 1200, -360, 80, 100, 310, 30)
        Case Is = 260: Call Galaga_Wave(2222, 8, 0, 5, 3000, -360, 80, 100, 310, 30)
        Case Is = 270: Call Galaga_Wave(1112, 8, 6, 9, 3000, -360, 80, 100, 310, 30)
        Case Is = 280: Call Galaga_Wave(2222, 8, 10, 14, 1000, -360, 80, 100, 310, 30)
        Case Is = 290: Call Galaga_Wave(2222, 8, 15, 19, 1000, -360, 80, 100, 310, 30)
        Case Is = 300: Call Galaga_Wave(2222, 8, 4, 7, 2600, -360, 80, 100, 310, 30)
        Case Is = 320: Call Galaga_Wave(2222, 8, 8, 11, 1000, -360, 85, 100, 310, 30)
        Case Is = 330: Call Galaga_Wave(2222, 8, 16, 19, 1500, -360, 85, 100, 310, 30)
        Case Is = 340: Call Galaga_Wave(2222, 8, 21, 25, 2000, -360, 80, 100, 310, 30)
        Case Is = 350: Call Galaga_Wave(2222, 8, 26, 30, 1100, -360, 80, 100, 310, 30)
        Case Is = 360: Call Galaga_Wave(2222, 8, 0, 4, 2000, -360, 80, 100, 310, 30)
        Case Is = 370: Call Galaga_Wave(2222, 8, 5, 8, 2900, -360, 80, 100, 310, 30)
        Case Is = 380: Call Galaga_Wave(2222, 8, 9, 14, 2965, -360, 80, 150, 315, 30)
        Case Is = 390: Call Galaga_Wave(2222, 8, 15, 18, 1040, -360, 80, 150, 315, 30)
        Case Is = 400: Call Galaga_Wave(2222, 8, 20, 23, 1000, -360, 80, 150, 315, 30)
        Case Is = 410: Call Galaga_Wave(2222, 8, 24, 28, 2000, -360, 80, 150, 315, 30)
        Case Is = 420: Call Galaga_Wave(2222, 8, 30, 34, 2800, -360, 80, 150, 315, 30)
        Case Is = 430: Call Galaga_Wave(2624, 6, 0, 5, 1200, -360, 80, 200, 315, 30)
        End Select
    End If
    
    If Ship_Stats(0) = 8 Then
        Select Case (Indic.Left - 120)
        Case Is = 20: Call Galaga_Wave(3332, 7, 0, 3, 1200, -360, 40, 250, 525, 15)
        Case Is = 40: Call Galaga_Wave(1112, 9, 4, 7, 2700, -360, 40, 250, 525, 15)
        Case Is = 60: Call Galaga_Wave(2222, 6, 8, 11, 1300, -360, 30, 300, 525, 15)
        Case Is = 80: Call Galaga_Wave(2624, 6, 12, 15, 2800, -360, 40, 300, 525, 15)
        Case Is = 100: Call Galaga_Wave(2426, 4, 16, 19, 2200, -360, 40, 300, 525, 15)
        Case Is = 120: Call Galaga_Wave(2222, 8, 20, 23, 2000, -360, 30, 300, 525, 15)
        Case Is = 130: Call Galaga_Wave(9999, 8, 24, 24, 2000, -360, 40, 800, 560, 50)
        Case Is = 140: Call Galaga_Wave(1112, 9, 4, 7, 2600, -360, 40, 250, 525, 15)
        Case Is = 160: Call Galaga_Wave(3212, 9, 8, 11, 2000, -360, 30, 400, 535, 15)
        Case Is = 180: Call Galaga_Wave(1232, 7, 12, 15, 1400, -360, 30, 400, 535, 15)
        Case Is = 200: Call Galaga_Wave(4444, 8, 0, 5, 3400, 2000, 30, 250, 525, 15)
        Case Is = 220: Call Galaga_Wave(6666, 8, 8, 13, 800, 2000, 30, 250, 525, 15)
        Case Is = 250: Call Galaga_Wave(3332, 7, 0, 3, 1200, -360, 40, 450, 525, 15)
        Case Is = 250: Call Galaga_Wave(1112, 9, 4, 7, 2700, -360, 40, 450, 535, 15)
        Case Is = 280: Call Galaga_Wave(1112, 9, 8, 11, 2700, -360, 40, 500, 535, 15)
        Case Is = 300: Call Galaga_Wave(3212, 9, 12, 15, 2600, -360, 30, 500, 535, 15)
        Case Is = 330: Call Galaga_Wave(2222, 8, 0, 3, 1855, -360, 40, 450, 535, 15)
        Case Is = 350: Call Galaga_Wave(2222, 6, 8, 11, 1300, -360, 40, 500, 455, 15)
        Case Is = 370: Call Galaga_Wave(2222, 6, 12, 15, 1300, -360, 40, 500, 545, 15)
        Case Is = 390: Call Galaga_Wave(2222, 6, 16, 19, 1300, -360, 40, 600, 545, 25)
        Case Is = 420: Call Galaga_Wave(6666, 4, 20, 23, 1040, -360, 40, 550, 525, 15)
        Case Is = 440: Call Galaga_Wave(4444, 6, 24, 27, 1000, -360, 30, 550, 525, 15)
        End Select
    End If
    
    If Ship_Stats(0) = 9 Then
        Select Case (Indic.Left - 120)
        Case Is = 30: Call Galaga_Wave(2624, 8, 0, 2, 1200, -360, 30, 500, 710, 10)
        Case Is = 50: Call Galaga_Wave(3322, 7, 3, 5, 1400, -360, 20, 500, 715, 15)
        Case Is = 70: Call Galaga_Wave(2426, 4, 6, 8, 2600, -360, 40, 500, 715, 15)
        Case Is = 90: Call Galaga_Wave(3332, 7, 17, 17, 1300, -360, 30, 550, 720, 10)
        Case Is = 110: Call Galaga_Wave(1112, 7, 18, 18, 3000, -360, 30, 550, 720, 10)
        Case Is = 130: Call Galaga_Wave(2222, 6, 0, 4, 1400, -360, 20, 550, 715, 10)
        Case Is = 170: Call Galaga_Wave(3212, 7, 9, 11, 1400, -360, 30, 500, 715, 10)
        Case Is = 190: Call Galaga_Wave(1222, 9, 12, 15, 2000, -360, 30, 500, 715, 10)
        Case Is = 210: Call Galaga_Wave(2222, 6, 5, 8, 1400, -360, 20, 500, 715, 10)
        Case Is = 230: Call Galaga_Wave(6666, 9, 0, 4, 100, 1800, 30, 550, 715, 10)
        Case Is = 250: Call Galaga_Wave(6666, 9, 8, 11, 100, 1800, 30, 500, 720, 15)
        Case Is = 270: Call Galaga_Wave(3332, 7, 17, 17, 1500, -360, 30, 550, 720, 10)
        Case Is = 290: Call Galaga_Wave(2624, 8, 17, 17, 1700, -360, 40, 600, 730, 15)
        Case Is = 310: Call Galaga_Wave(2624, 8, 18, 18, 1900, -360, 40, 600, 730, 15)
        Case Is = 330: Call Galaga_Wave(2222, 8, 6, 8, 2365, -360, 40, 750, 725, 15)
        Case Is = 350: Call Galaga_Wave(2222, 8, 8, 11, 1040, -360, 40, 750, 725, 15)
        Case Is = 370: Call Galaga_Wave(2222, 8, 0, 3, 1700, -360, 30, 750, 725, 15)
        Case Is = 390: Call Galaga_Wave(2222, 6, 0, 4, 1400, -360, 20, 750, 715, 10)
        Case Is = 410: Call Galaga_Wave(2624, 8, 0, 2, 1200, -360, 30, 700, 700, 10)
        Case Is = 430: Call Galaga_Wave(3322, 7, 3, 5, 1400, -360, 20, 700, 750, 15)
        End Select
    End If
    
    If Ship_Stats(0) = 10 Then
        Select Case (Indic.Left - 120)
        Case Is = 60: Call Galaga_Wave(9999, 8, 32, 32, 1500, -200, 35, 1000, 1605, 100)
        Case Is = 130: Call Galaga_Wave(9999, 8, 33, 33, 1600, -360, 35, 2000, 1685, 100)
        Case Is = 200: Call Galaga_Wave(9999, 8, 34, 34, 2800, -360, 35, 3000, 1635, 100)
        Case Is = 280: Call Galaga_Wave(9999, 8, 35, 35, 1800, -300, 35, 4000, 1735, 100)
        Case Is = 360: Call Galaga_Wave(9999, 8, 36, 36, 1540, -260, 35, 5000, 1735, 100)
        Case Is = 400: Call Galaga_Wave(9999, 8, 37, 37, 2800, -360, 35, 6000, 1735, 100)
        
        Case Is = 460: Call Galaga_Wave(2222, 6, 0, 3, 1200, 1700, 0, 10000, 3650, 5)
        Case Is = 460: Call Galaga_Wave(2222, 6, 4, 7, 1400, 1750, 0, 10000, 3650, 5)
        Case Is = 460: Call Galaga_Wave(2222, 6, 8, 11, 1800, 1700, 0, 10000, 3650, 5)
        Case Is = 460: Call Galaga_Wave(2222, 6, 12, 15, 2000, 1750, 0, 10000, 3650, 5)
        End Select
    End If

End Sub

Private Sub Galaga_Wave(g_Move_pattern, g_Position_dir, g_Number, g_Number_End, g_Pos_x, g_Pos_y, g_Speed, g_Health, g_worth, g_Damage) 'Create New Wave
        Galaga_Move_Pat_Inv = g_Move_pattern
        Galaga_Position_Pat = g_Position_dir
        Galaga_Start = g_Number
        Galaga_Num = g_Number_End
        Galaga_Start_POSx = g_Pos_x
        Galaga_Start_POSy = g_Pos_y
        Galaga_Speed = g_Speed
        Galaga_Init_Life = g_Health
        Galaga_Worth_Inv = g_worth
        Galaga_Dam_Inv = g_Damage
        
        If BossAppear = True Then Galaga_Boss_Bar = Galaga_Init_Life
                
        Call Summon_Galaga
End Sub
Private Sub Summon_Galaga()     'Draws out galaga
        For x = Galaga_Start To Galaga_Num
            Galaga_Dam(x) = Galaga_Dam_Inv
            Galaga_Worth(x) = Galaga_Worth_Inv
            Galaga_Move_Pat(x) = Galaga_Move_Pat_Inv
            Galaga_Speedx(x) = Galaga_Speed
            Galaga_Life(x) = Galaga_Init_Life
            Galaga(x).Top = Galaga_Start_POSy
            Galaga(x).Left = Galaga_Start_POSx
            Galaga(x).Visible = True
                If Galaga_Position_Pat = 4 Then Galaga_Start_POSx = Galaga_Start_POSx - 360
                If Galaga_Position_Pat = 6 Then Galaga_Start_POSx = Galaga_Start_POSx + 360
                If Galaga_Position_Pat = 8 Then Galaga_Start_POSy = Galaga_Start_POSy - 360
                If Galaga_Position_Pat = 7 Then
                    Galaga_Start_POSx = Galaga_Start_POSx - 360
                    Galaga_Start_POSy = Galaga_Start_POSy - 360
                End If
                If Galaga_Position_Pat = 9 Then
                    Galaga_Start_POSx = Galaga_Start_POSx + 360
                    Galaga_Start_POSy = Galaga_Start_POSy - 360
                End If
            Galaga_Start = Galaga_Start + 1
        Next x
End Sub


Private Sub GalagaMove_Timer()

'------------------MOVEMENT PATTERNS----------------------
    For x = 0 To 49
        
        If Galaga(x).Visible = True Then
            
            '------2624------
            If Galaga_Move_Pat(x) = 2624 Then
                If Galaga(x).Top <= 400 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
                If Galaga(x).Top > 400 And Galaga(x).Top < 500 Then
                    Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                    End If
                If Galaga(x).Left >= 2700 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left
                    End If
                If Galaga(x).Top >= 1800 And Galaga(x).Left > 1000 Then Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                If Galaga(x).Left <= 1000 Then Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
            End If
        
            '------2426------
            If Galaga_Move_Pat(x) = 2426 Then
                If Galaga(x).Top <= 400 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
                If Galaga(x).Top > 400 And Galaga(x).Top < 500 Then
                    Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                    End If
                If Galaga(x).Left <= 1100 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left
                    End If
                If Galaga(x).Top >= 1800 And Galaga(x).Left < 1100 Then Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                If Galaga(x).Left >= 2000 Then Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
            End If
        
            '------2222------
            If Galaga_Move_Pat(x) = 2222 Then
                Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
            End If
        
            '------4444------
            If Galaga_Move_Pat(x) = 4444 Then
                Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                If Galaga(x).Left <= 700 Then Galaga(x).Visible = False
            End If
        
            '------6666------
            If Galaga_Move_Pat(x) = 6666 Then
                Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                If Galaga(x).Left >= 3400 Then Galaga(x).Visible = False
            End If
        
            '------3322------
            If Galaga_Move_Pat(x) = 3322 Then
                If Galaga(x).Left <= 2300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                    End If
                If Galaga(x).Left >= 2300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
        
            '------1122------
            If Galaga_Move_Pat(x) = 1122 Then
                If Galaga(x).Left > 1300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                    End If
                If Galaga(x).Left <= 1300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
        
            '------3332------
            If Galaga_Move_Pat(x) = 3332 Then
                If Galaga(x).Left <= 2850 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                    End If
                If Galaga(x).Left >= 2850 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
        
            '------1112------
            If Galaga_Move_Pat(x) = 1112 Then
                If Galaga(x).Left > 1070 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                    End If
                If Galaga(x).Left <= 1070 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
            
            '------3212------
            If Galaga_Move_Pat(x) = 3212 Then
                If Galaga(x).Left <= 3450 Or Galaga(x).Top < 760 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                    End If
                If Galaga(x).Left > 2750 Or Galaga(x).Top <= 1300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
                If Galaga(x).Top > 1300 Or Galaga(x).Left > 1100 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                    End If
                If Galaga(x).Left <= 1100 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
        
            '------1232------
            If Galaga_Move_Pat(x) = 1232 Then
                If Galaga(x).Left >= 850 Or Galaga(x).Top < 760 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                    End If
                If Galaga(x).Left < 1250 Or Galaga(x).Top <= 1300 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
                If Galaga(x).Top > 1300 Or Galaga(x).Left < 2600 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                    End If
                If Galaga(x).Left <= 2600 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
            End If
        
            '------9999------
            If Galaga_Move_Pat(x) = 9999 Then
                If Galaga(x).Top <= 360 Then
                    Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                    End If
                If Galaga(x).Top > 360 Then
                    If Galaga(x).Top < 550 And (Galaga(x).Left + Galaga(x).Width) < 3265 Then
                            Galaga(x).Left = Galaga(x).Left + Galaga_Speedx(x)
                        End If
                    If (Galaga(x).Left + Galaga(x).Width) >= 3265 Then
                        Galaga(x).Top = Galaga(x).Top + Galaga_Speedx(x)
                        End If
                    If Galaga(x).Top >= 910 And Galaga(x).Left > 820 Then
                        Galaga(x).Left = Galaga(x).Left - Galaga_Speedx(x)
                        End If
                    If Galaga(x).Left <= 820 Then
                        Galaga(x).Top = Galaga(x).Top - Galaga_Speedx(x)
                        End If
                End If
            End If
        
        
        
        End If
    If Galaga(x).Top >= 2900 Then Galaga(x).Visible = False
    If Galaga(x).Left <= 0 Then Galaga(x).Visible = False
    If Galaga(x).Left >= 4000 Then Galaga(x).Visible = False
    Next x

End Sub

Private Sub StatusMove_Timer()  'Buffer texts
    For x = 0 To 2
    lblStatus(x).Caption = lblStatus(x + 1).Caption
    Next x
    lblStatus(3).Caption = ""
End Sub

Private Sub Send_Msg(Message As String) 'Buffer control
    For x = 0 To 3
        If lblStatus(x) <> "" Then
        Else:
            lblStatus(x).Caption = Message
            x = 3
        End If
    Next x
End Sub
