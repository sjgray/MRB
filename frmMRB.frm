VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMRB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-ROM Builder"
   ClientHeight    =   8730
   ClientLeft      =   240
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9870
   Begin VB.CheckBox cbMode16 
      Caption         =   "16 Bit Mode"
      Height          =   255
      Left            =   8400
      TabIndex        =   101
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "S&plit"
      Height          =   405
      Left            =   6630
      TabIndex        =   100
      Top             =   8280
      Width           =   1305
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Set"
      Height          =   345
      Left            =   8760
      TabIndex        =   10
      Top             =   750
      Width           =   1065
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   5670
      TabIndex        =   16
      Top             =   8280
      Width           =   765
   End
   Begin VB.CheckBox cbAllowEmpty 
      Caption         =   "Allow empty slots"
      Height          =   255
      Left            =   6510
      TabIndex        =   5
      Top             =   1410
      Width           =   1545
   End
   Begin VB.ComboBox cboGroup 
      Height          =   315
      ItemData        =   "frmMRB.frx":0000
      Left            =   5310
      List            =   "frmMRB.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1065
   End
   Begin VB.TextBox txtPad 
      Height          =   285
      Left            =   9300
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Enter a decimal value"
      Top             =   1110
      Width           =   465
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6225
      LargeChange     =   4
      Left            =   9450
      TabIndex        =   72
      Top             =   1980
      Width           =   375
   End
   Begin VB.ComboBox cboNumSlots 
      Height          =   315
      ItemData        =   "frmMRB.frx":0004
      Left            =   3360
      List            =   "frmMRB.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   945
   End
   Begin VB.CheckBox cbAllowShort 
      Caption         =   "Allow short files"
      Height          =   255
      Left            =   6510
      TabIndex        =   4
      Top             =   1140
      Width           =   1545
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   720
      Width           =   5205
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move &DOWN"
      Height          =   405
      Left            =   2790
      TabIndex        =   13
      Top             =   8280
      Width           =   1125
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move &UP"
      Height          =   405
      Left            =   1650
      TabIndex        =   12
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdIns 
      Caption         =   "&Insert"
      Height          =   405
      Left            =   4860
      TabIndex        =   15
      Top             =   8280
      Width           =   765
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7920
      Top             =   8250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   15
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   33
      Top             =   7860
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   14
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   32
      Top             =   7470
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   13
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   31
      Top             =   7080
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   12
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   30
      Top             =   6690
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   11
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   29
      Top             =   6300
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   10
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   28
      Top             =   5910
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   9
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   27
      Top             =   5520
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   8
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   5130
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   7
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   25
      Top             =   4740
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   6
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   4350
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   5
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   23
      Top             =   3960
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   4
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Top             =   3570
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   3
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      Top             =   3180
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   2
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   2790
      Width           =   6855
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   1
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   2400
      Width           =   6855
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build It!"
      Height          =   435
      Left            =   8430
      TabIndex        =   17
      Top             =   8280
      Width           =   1425
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "D&elete"
      Height          =   405
      Left            =   4050
      TabIndex        =   14
      Top             =   8280
      Width           =   765
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add &File..."
      Height          =   405
      Left            =   90
      TabIndex        =   11
      Top             =   8280
      Width           =   1425
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   0
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   2010
      Width           =   6855
   End
   Begin VB.CommandButton cmdSaveSet 
      Caption         =   "&Save Set"
      Height          =   345
      Left            =   7620
      TabIndex        =   9
      Top             =   750
      Width           =   1065
   End
   Begin VB.CommandButton cmdLoadSet 
      Caption         =   "&Load Set"
      Height          =   345
      Left            =   6480
      TabIndex        =   8
      Top             =   750
      Width           =   1065
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   645
      Left            =   8760
      TabIndex        =   7
      Top             =   60
      Width           =   1065
   End
   Begin VB.ComboBox cboTargetSize 
      Height          =   315
      ItemData        =   "frmMRB.frx":0008
      Left            =   1170
      List            =   "frmMRB.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1470
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "< Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8460
      TabIndex        =   99
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Slot Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   98
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4530
      TabIndex        =   97
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Cmd"
      Height          =   195
      Left            =   8040
      TabIndex        =   96
      ToolTipText     =   "Is a Command not a File"
      Top             =   1770
      Width           =   315
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000016&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      Height          =   195
      Left            =   7830
      Top             =   1770
      Width           =   165
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Pad (0-255):"
      Height          =   195
      Left            =   8400
      TabIndex        =   95
      Top             =   1170
      Width           =   870
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7020
      Top             =   1770
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "No File"
      Height          =   195
      Left            =   7230
      TabIndex        =   94
      ToolTipText     =   "The file does not exist"
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Error"
      Height          =   195
      Left            =   6540
      TabIndex        =   93
      ToolTipText     =   "The file is empty, too short, or too big."
      Top             =   1770
      Width           =   330
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Warn"
      Height          =   195
      Left            =   5820
      TabIndex        =   92
      ToolTipText     =   "The file appears to have 2 extra Load Address bytes"
      Top             =   1770
      Width           =   390
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   8460
      TabIndex        =   91
      Top             =   2430
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   15
      Left            =   8460
      TabIndex        =   90
      Top             =   7890
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   14
      Left            =   8460
      TabIndex        =   89
      Top             =   7500
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   13
      Left            =   8460
      TabIndex        =   88
      Top             =   7110
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   12
      Left            =   8460
      TabIndex        =   87
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   11
      Left            =   8460
      TabIndex        =   86
      Top             =   6330
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   8460
      TabIndex        =   85
      Top             =   5940
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   8460
      TabIndex        =   84
      Top             =   5550
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   8460
      TabIndex        =   83
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   8460
      TabIndex        =   82
      Top             =   4770
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   8460
      TabIndex        =   81
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   8460
      TabIndex        =   80
      Top             =   3990
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   8460
      TabIndex        =   79
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   8460
      TabIndex        =   78
      Top             =   3210
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   8460
      TabIndex        =   77
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00808080&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   8460
      TabIndex        =   76
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   75
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   74
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   73
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label lblCalc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3870
      TabIndex        =   71
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Slots:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   70
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Target:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   69
      Top             =   1080
      Width           =   750
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   6300
      Top             =   1770
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4860
      Top             =   1770
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   5610
      Top             =   1770
      Width           =   165
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   630
      TabIndex        =   67
      Top             =   7860
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   630
      TabIndex        =   66
      Top             =   7470
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   630
      TabIndex        =   65
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   630
      TabIndex        =   64
      Top             =   6690
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   630
      TabIndex        =   63
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   630
      TabIndex        =   62
      Top             =   5910
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   630
      TabIndex        =   61
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   630
      TabIndex        =   60
      Top             =   5130
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   630
      TabIndex        =   59
      Top             =   4740
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   630
      TabIndex        =   58
      Top             =   4350
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   630
      TabIndex        =   57
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   630
      TabIndex        =   56
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   630
      TabIndex        =   55
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   630
      TabIndex        =   54
      Top             =   2790
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   630
      TabIndex        =   53
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   630
      TabIndex        =   52
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Set Desc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   51
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   60
      TabIndex        =   50
      Top             =   7830
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   60
      TabIndex        =   49
      Top             =   7440
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   60
      TabIndex        =   48
      Top             =   7050
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   60
      TabIndex        =   47
      Top             =   6660
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   60
      TabIndex        =   46
      Top             =   6270
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   60
      TabIndex        =   45
      Top             =   5880
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   60
      TabIndex        =   44
      Top             =   5490
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   60
      TabIndex        =   43
      Top             =   5100
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   60
      TabIndex        =   42
      Top             =   4740
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   60
      TabIndex        =   41
      Top             =   4350
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   60
      TabIndex        =   40
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   60
      TabIndex        =   39
      Top             =   3570
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   60
      TabIndex        =   38
      Top             =   3180
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   60
      TabIndex        =   37
      Top             =   2790
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   60
      TabIndex        =   36
      Top             =   2400
      Width           =   525
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   35
      Top             =   2010
      Width           =   525
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmMRB.frx":000C
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   60
      Width           =   8595
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Good"
      Height          =   195
      Left            =   5070
      TabIndex        =   68
      ToolTipText     =   "File is valid"
      Top             =   1770
      Width           =   390
   End
End
Attribute VB_Name = "frmMRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MRB - Multi-ROM Builder, (C) 2024 Steve J. Gray
' ===
' This is a utility to assemble binary files to be written to ROMS/EPROMS etc of various sizes that
' can be specified, as well as how many "slots" it will contain. For each slot you can specify a
' file that will be loaded into the slots, or you can specify a COMMAND that can fill the slot.
' The program will verify file sizes for each slots and indicate which is valid using a status
' colour. Slots can be re-arranged, deleted, inserted or cleared etc. Drag and Drop of files into
' slots is supported. You can specify if short files or empty slots are accepted. Empty slots can
' be filled with a specified byte value.

Dim SelNum        As Integer        'Selected Set number (0-255)
Dim SelIndex      As Integer        'Selected Index      (0-15)
Dim EdNum         As Integer        'Current Edit number (0-255)
Dim EdIndex       As Integer        'Current Edit Index  (0-15)
Dim TopNum        As Integer        'Top Slot Number     (0-255)

Dim TargetSize    As Single         'Binary size
Dim SlotSize      As Single         'Size of each slot (bytes)
Dim NumSlots      As Integer        'Number of slots (2-64)

Dim MaxSlot       As Integer        'Max Slot number

Dim Cr            As String         'Carriage Return
Dim Ready         As Boolean        'Flag to indicate init complete

Dim File(255)     As String         'Filename for slot including path
Dim Base(255)     As String         'Base filename for display
Dim SlotAddr(255) As String         'Address offset - to speed up scrolling - calculated
Dim FileInfo(255) As String         'File Info (size/CMD)
Dim FileSize(255) As Single         'FileSize (bytes)

Dim CMD           As String         'CMD string indicator
Dim GroupFlag     As Boolean        'Flag to indicate grouping enabled
Dim GroupSize     As Integer        'Group size

'---- Show the Program About Box
Private Sub cmdAbout_Click()
    MsgBox "Multi-ROM Builder, (C)2024 Steve J. Gray" & Cr & "Version 1.2 - Jun 1/2024"
End Sub

'---- Start Here
Private Sub Form_Load()
    Dim A As Integer

    Ready = False
    Cr = Chr(13)            'Carriage Return
    CMD = "%CMD"            'Command indicator string
    
    With cboTargetSize
        .Clear
        .AddItem "8K (2764)"
        .AddItem "16K (27128)"
        .AddItem "32K (27256)"
        .AddItem "64 K (27512)"
        .AddItem "128K (27C010)"
        .AddItem "256K (27C020)"
        .AddItem "512K (27C040)"
        .AddItem "1MB (27C080)"
        .AddItem "2MB (27C100)"
        .ListIndex = 3
    End With
    
    With cboNumSlots
        .Clear
        .AddItem "2"
        .AddItem "4"
        .AddItem "8"
        .AddItem "16"
        .AddItem "32"
        .AddItem "64"
        .AddItem "128"
        .AddItem "256"
        .ListIndex = 3
    End With
    
    With cboGroup
        .Clear
        .AddItem "None"
        .AddItem "2"
        .AddItem "4"
        .AddItem "8"
        .AddItem "16"
        .AddItem "32"
        .ListIndex = 0
    End With
        
    For A = 0 To 15
        lblN(A).ForeColor = vbWhite
        lblK(A).BackColor = vbGreen
        lblSize(A).ForeColor = vbBlack
    Next A
    
    NewSet                'Init Slots
    
    Ready = True            'Init complete
    SetSlotSize             'Init Slot size and Draw the slots (hides as needed, uses TopNum)
    
End Sub

'---- Set Slot Size
Private Sub SetSlotSize()
    Dim A As Integer
    
    TargetSize = 2 ^ (cboTargetSize.ListIndex + 3) * 1024
    NumSlots = 2 ^ (cboNumSlots.ListIndex + 1)
    SlotSize = TargetSize / NumSlots
    MaxSlot = NumSlots - 1
        
    GroupSize = 2 ^ cboGroup.ListIndex
    GroupFlag = False: If GroupSize > 1 Then GroupFlag = True
    
    lblCalc.Caption = Str(SlotSize) & " bytes."
    
    If NumSlots <= 16 Then
        VScroll1.Visible = False
        TopNum = 0
    Else
        VScroll1.Min = 0
        VScroll1.Max = NumSlots - 16
        VScroll1.Visible = True
    End If
    
    For A = 0 To MaxSlot
        SlotAddr(A) = Hex(A * SlotSize)
    Next A
    
    DrawSlots
End Sub

'--- Draw/Update Slot entries
' This will take the TopNum variable and fill in the info for all visible slots,
' update the slot number/colour, address offset, filename, and set text and background
' colour (according to filesize comparison with SLOTSIZE).
    
Private Sub DrawSlots()
    Dim A As Integer, C As Integer, N As Integer
    Dim FIO As Integer, FLen As Single
    Dim ShortFlag As Boolean, EmptyFlag As Boolean
    Dim Filename As String
    
    If Ready = False Then Exit Sub
          
    ShortFlag = False: If cbAllowShort.Value = vbChecked Then ShortFlag = True
    EmptyFlag = False: If cbAllowEmpty.Value = vbChecked Then EmptyFlag = True
    
    'Debug.Print "DrawSlots: TopNum="; TopNum; " EdIndex="; EdIndex
    
    For A = 0 To 15
        If A < NumSlots Then
            N = TopNum + A                             'N is the pointer to the actual slot
            Filename = File(N)                          'Get the Filename/Cmd string
            
            If N = SelNum Then
                lblN(A).BackColor = vbRed               'Selected is made RED
            Else
                lblN(A).BackColor = vbBlue              'Un-Selected is BLUE
            End If
            
            lblN(A).Caption = Str(N)                    'Slot Number
            lblK(A).Caption = SlotAddr(N)               'Address Offset
            txtFN(A).Text = Base(N)                     'Filename entry
            
            If (GroupFlag = True) And (N Mod GroupSize) = 0 Then
                    lblK(A).BackColor = vbGreen         'Light Green means start of group
            Else
                lblK(A).BackColor = &HC000&             'Dk.Green normal
            End If
                  
            If Left(Filename, 1) = "%" Then
                lblSize(A).Caption = CMD                'COMMAND identifier
                lblSize(A).BackColor = &HE0E0E0         'Lt.Grey
            Else
                If FileInfo(N) = "?" Then               'Check if file exists (ie new set loaded)
                    If Exists(Filename) = True Then
                        FLen = FileLen(Filename)        'Get the length then close it
                        FileInfo(N) = Str(FLen)         'Set Info and Length
                        FileSize(N) = FLen
                    Else
                        FileInfo(N) = ""
                        FileSize(N) = 0
                    End If
                End If

                lblSize(A).Caption = FileInfo(N)        'File Info/Size
                FLen = FileSize(N)                      'Check File size against Slotsize
                If FLen = 0 Then
                    If EmptyFlag = True Then
                        lblSize(A).BackColor = vbBlack  'Black if 0 and Empty allowed
                    Else
                        lblSize(A).BackColor = vbRed    'Black if 0
                    End If
                    
                ElseIf FLen < SlotSize Then
                    If ShortFlag = True Then
                        lblSize(A).BackColor = &HC000&      'Green if Good and short files allowed
                    Else
                        lblSize(A).BackColor = vbRed    'Otherwise Red
                    End If
                    
                ElseIf FLen = SlotSize Then
                        lblSize(A).BackColor = &HC000&      'Green if Good
                        
                ElseIf FLen - 2 = SlotSize Then
                    lblSize(A).BackColor = vbYellow     'Yellow if Load Address included
                    
                Else
                    lblSize(A).BackColor = vbRed        'Red if Greater
                End If
            End If
                      
            lblN(A).Visible = True                  'Slot Number
            lblK(A).Visible = True                  'Address Offset
            txtFN(A).Visible = True                 'Filename
            lblSize(A).Visible = True               'File Size
        Else
            lblN(A).Visible = False
            txtFN(A).Visible = False
            lblK(A).Visible = False
            lblSize(A).Visible = False
        End If
        
    Next A
    
    If VScroll1.Visible = True Then VScroll1.Value = TopNum     'Update scrollbar if visible
    
    DoEvents
    
End Sub


Private Sub SplitSlot()

    Dim Filename As String, OutFile As String, OutFile2 As String
    Dim FIO As Integer, FIO2 As Integer, FIO3 As Integer
    Dim Temp As String, BY As String * 1
    Dim Mode16 As Boolean
    
    Dim C As Integer, i As Single, j As Single, SS As Integer
            
    Mode16 = False: SS = 1
    If cbMode16.Value = vbChecked Then Mode16 = True: SS = 2
    
    Filename = File(0)
    
    If Filename = "" Or Left(Filename, 1) = "%" Then MsgBox "Please enter a valid filename in SLOT 0!": Exit Sub
    If Exists(Filename) = False Then MsgBox "The file does not exist!": Exit Sub
    If FileLen(Filename) <> TargetSize Then MsgBox "The filesize does not match Target size (" & Str(TargetSize) & " bytes)!": Exit Sub
    
    Temp = "This will split the FILE in SLOT 0." & Cr & "Make sure the ROM size and SlotSize are correct!"
    If Mode16 = True Then Temp = Temp & Cr & "16-bit mode ENABLED!:" & Cr & "Source byte pairs are split into EVEN/ODD output files."
    
    If MsgBox(Temp, vbOKCancel, "Split SLOT 0 file.") = vbCancel Then Exit Sub
    
    Temp = FNoExt(Filename)                                          'The file path without extension
    
    FIO = FreeFile
    Open Filename For Binary As FIO                                 'Open the source file to READ
    
    For i = 0 To MaxSlot Step SS                                    'Loop for each slot
        OutFile = Temp & "." & Format(i, "000")                      'Build the output filename
        FIO2 = FreeFile
        Open OutFile For Output As FIO2                             'Open destination file to WRITE
        File(i) = OutFile                                           'Filename with path
        Base(i) = FName(OutFile)                                    'Filename only
        FileSize(i) = SlotSize                                      'Filesize same as slotsize
        FileInfo(i) = Str(SlotSize)                                 'FileInfo
        
        If Mode16 = True Then
            OutFile2 = Temp & "." & Format(i + 1, "000")             'Build the output filename
            FIO3 = FreeFile
            Open OutFile2 For Output As FIO3                        'Open destination file to WRITE
            File(i + 1) = OutFile2                                  'Filename with path
            Base(i + 1) = FName(OutFile2)                           'Filename only
            FileSize(i + 1) = SlotSize                              'Filesize same as slotsize
            FileInfo(i + 1) = Str(SlotSize)                         'FileInfo
        End If
                
        
        DoEvents
        For j = 1 To SlotSize Step SS                               'Loop for slotsize
            BY = Input(FIO, 1)                                      'Read from source
            Print #FIO2, BY;                                        'Write to output
            
            If Mode16 = True Then
                BY = Input(FIO, 1)                                  'Read from source
                Print #FIO3, BY;                                    'Write to output
            End If
        Next j
        Close FIO2
        If Mode16 = True Then Close FIO3
    Next i
    
    Close FIO
    MsgBox "File has been split!"
    
    DrawSlots
    
End Sub

'---- UpdateEdit
' When a SLOT is selected and being edited, any operation on that SLOT by other
' GUI elements must ensure that the SLOT data is update BEFORE it is operated on.
' If the filename starts with "%" then it is a command, so it is copied to the BASE name.
Private Sub UpdateEdit()
    If EdNum < 0 Then
        'Debug.Print "UpdateEdit abort"
        Exit Sub                  'Exit if NO slot selected
    End If
    
    'Debug.Print "UpdateEdit"; EdNum
    
    File(EdNum) = txtFN(EdIndex).Text           'Restore updated filename
    If Left(File(Edmum), 1) = "%" Then
        Base(EdNum) = File(EdNum)               'Copy CMD to BASE name
    Else
        Base(EdNum) = FName(File(EdNum))        'Update BASE name from current filename
    End If
    EdNum = -1
End Sub

'==== SLOT HANDLING

'---- Select via Address Offset box
Private Sub lblK_Click(Index As Integer)
    SelectN Index
End Sub

'---- Select via Index Box
Private Sub lblN_Click(Index As Integer)
    SelectN Index
End Sub

'---- Double-click on Index button to load a binary
Private Sub lblN_DblClick(Index As Integer)
    cmdAdd_Click
End Sub


'---- Select the Text edit box
Private Sub txtFN_GotFocus(Index As Integer)
    SelectN Index
End Sub

'---- Select a specific Text Box
Private Sub SelectN(ByVal Index As Integer)
    'Debug.Print "SelectN"; Index
    
    '-- Check for scrolling
    If (Index < 0) Then
        If TopNum > 0 Then TopNum = TopNum - 1          'We can scroll UP
    End If
    
    If (Index > 15) Then
        If MaxSlot > 15 Then
            If (TopNum + Index) < MaxSlot Then
                TopNum = TopNum + 1: Index = Index - 1  'We can scroll DOWN
            Else
                TopNum = MaxSlot - 15                   'We can't scroll - Force to last position
            End If
        End If
    End If
    
    '-- Ensure Index and TopNum are in range
    If MaxSlot < 16 Then TopNum = 0
    If Index < 0 Then Index = 0
    If Index > 15 Then Index = 15
    If Index > MaxSlot Then Index = MaxSlot
    
    '-- Set Num and Index
    SelNum = TopNum + Index                         'Index to slot data
    SelIndex = Index
    EdNum = SelNum
    EdIndex = Index                                 'Visible Slot number
    
    DrawSlots
    
    txtFN(EdIndex).Text = File(EdNum)               'Edit the full path
    txtFN(EdIndex).SetFocus                         'Activate the cursor

    'Debug.Print "SelectN done"
End Sub

Private Sub txtFN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim N As Integer
    
    Select Case KeyCode
        Case 36 'Home
            TopNum = 0
            N = -15: GoSub KeyDSub
            
        Case 38 'Cursor Up
            N = -1: GoSub KeyDSub
            
        Case 40 'Cursor Down
            N = 1: GoSub KeyDSub
            
    End Select
    Exit Sub
    
KeyDSub:
    UpdateEdit
    EdIndex = EdIndex + N: SelectN EdIndex
    KeyCode = 0
    Return
    
End Sub

'---- Handle Keystrokes in text boxes
Private Sub txtFN_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If (Index < 15) And (Index < SlotSize) Then
                txtFN(Index + 1).SetFocus
                KeyAscii = 0
            End If
            
            If (Index = 15) And (TopNum + 15) < MaxSlot Then
                TopNum = TopNum + 1                       'Move top slot down
                SelNum = SelNum + 1                         'Make new slot selected
                Index = 0
                DrawSlots                                   'Draw all slots
                txtFN(15).SetFocus                          'Put cursor in edit box
            End If
        Case 31 To 127
        Case Else
            'Debug.Print "KeyAscii="; KeyAscii
    End Select
End Sub

Private Sub VScroll1_Change()
    TopNum = VScroll1.Value
    DrawSlots
End Sub

Private Sub VScroll1_Scroll()
    TopNum = VScroll1.Value
    DrawSlots
End Sub

'==== SUBROUTINES

Private Sub NewSet()
    Dim A As Integer
    
    For A = 0 To 255
        File(A) = "%FILL" & Str(A)
        Base(A) = File(A)
        FileSize(A) = 0
        FileInfo(A) = CMD
    Next A

    TopNum = 0                      'Set to top of list
    SelNum = 0                      'Selected SET Number
    SelIndex = 0                    'Selected visible SLOT
    EdNum = -1                      'Edit SET number (-1 means none set)
    EdIndex = 0                     'Edit visible SLOT
    
    txtDesc.Text = "Multi-ROM Set"
    DoEvents
End Sub
Private Sub AddBin()
    Dim Filename As String
    
    Filename = FileOpenSave("", 0, 2, "Add ROM")
    If Filename <> "" Then
        File(SelNum) = Filename
        Base(SelNum) = FName(Filename)
        txtFN(SelIndex).Text = Filename: DoEvents
        FileInfo(SelNum) = "?"
        FileSize(SelNum) = 0
        
        SelectN SelIndex
    End If

End Sub

Private Sub LoadSet()
    Dim Filename As String
    Dim FIO As Integer, i As Integer, Tmp As String
    
    On Local Error Resume Next                          'Allow incomplete set file
    
    Filename = FileOpenSave("", 0, 1, "Load Set")
    If Exists(Filename) = True Then
    
        FIO = FreeFile
        Open Filename For Input As FIO
        i = 0
        Line Input #FIO, Tmp: txtDesc.Text = Tmp        'Set Description
        Do While Not EOF(FIO)
            Tmp = ""
            Line Input #FIO, Tmp                        'Path+Filename
            File(i) = Tmp                               'Filename including path
            Base(i) = FName(Tmp)                        'Base filename - no path (for display)
            FileInfo(i) = "?"                           'File Size string
            FileSize(i) = 0                             'File Size
            
            i = i + 1: If i > 255 Then Exit Do
        Loop
        
        Close FIO
        SelectN 0                                       'Select first file slot
    End If
    
End Sub

Private Sub SaveSet()
    Dim Filename As String
    Dim FIO As Integer, i As Integer, Tmp As String
    
    Filename = FileOpenSave("", 1, 1, "Save Set"): If Filename = "" Then Exit Sub
    If Overwrite(Filename) = True Then
        FIO = FreeFile
        Open Filename For Output As FIO
        Print #FIO, txtDesc.Text                    'Set Description
        For i = 0 To MaxSlot
            Print #FIO, File(i)                     'Path+Filename
        Next i
        Close FIO
    End If
    
End Sub

'==================
' Slot
'==================

'---- SwapSlots
' Swaps two specified slots - all 4 data
Private Sub SwapSlots(ByVal P1 As Integer, ByVal P2 As Integer)
    Dim Tmp As String, N As Single
    
        Tmp = Base(P1): Base(P1) = Base(P2): Base(P2) = Tmp
        Tmp = File(P1): File(P1) = File(P2): File(P2) = Tmp
        Tmp = FileInfo(P1): FileInfo(P1) = FileInfo(P2): FileInfo(P2) = Tmp
        N = FileSize(P1): FileSize(P1) = FileSize(P2): FileSize(P2) = N
End Sub

'---- MoveDown
' Moves the selected slot down - slot below is moved up
Private Sub MoveDown()
    Dim N As Integer
    
    N = TopNum + SelIndex
    'Debug.Print "--MoveDownSub: TopNum="; TopNum; " SelNum="; N: DoEvents
    If N < MaxSlot Then
        SwapSlots N, N + 1
        'SelectN SelNum + 1
    End If
    'Debug.Print "MoveDownSub done."; N
End Sub

'---- MoveUp
' Moves the selected slot up - slot above moves down
Private Sub MoveUp()
    Dim N As Integer
    
    N = TopNum + SelNum
    
    If SelNum > 0 Then
        SwapSlots N, N - 1
        'SelectN SelNum - 1
    End If
    
End Sub

'---- DelSlot
' Deletes the selected slot - all slots below move up
Private Sub DelSlot()
    Dim N As Integer, i As Integer
    
    N = TopNum + SelNum

    If N < MaxSlot Then
        For i = N To MaxSlot - 1
            SwapSlots i, i + 1
        Next
    End If
    
    ClearSlot MaxSlot
    DrawSlots
End Sub

'---- InsSlot
' Moves the selected slot and all slots below down then clears the selected slot
Private Sub InsSlot()
    Dim N As Integer, i As Integer
    
    N = TopNum + SelNum

    If N < MaxSlot Then
        For i = MaxSlot To N + 1 Step -1
            SwapSlots i, i - 1
        Next
    End If
    
    ClearSlot N
    DrawSlots
        
End Sub

'---- ClearSlot
' Erases all data for the current slot
Private Sub ClearSlot(ByVal N As Integer)
    File(N) = ""
    Base(N) = ""
    FileInfo(N) = ""
    FileSize(N) = 0
End Sub

'=====================================================================
' BUILD IT!
'=====================================================================
' Verifies all slots are valid using selected options.
' If not valid displays message on first bad slot.
' If all are valid, gets Target filename then creates Target file.
Private Sub cmdBuild_Click()
    Dim Filename As String, Filename2 As String
    Dim Padd As String * 1, DefPadd As String * 1, CMD As String * 1
    Dim B As String * 1, LA As String * 2, PARM As String
    Dim FIO As Integer, FIO2 As Integer, FIO3 As Integer
    Dim FLen As Single, C As Single
    Dim i As Integer, j As Integer, p As Integer
    Dim Mode As Integer, Mode16 As Boolean, SS As Integer
    

    j = Val(txtPad.Text): If j < 0 Or j > 255 Then j = 0
    DefPadd = Chr(j)
    
    Mode16 = False: SS = 1
    If cbMode16.Value = vbChecked Then Mode16 = True: SS = 2
    
    '--- Check all slots for files exist, or CMD, and filesize is in range as appropriate
    
    For i = 0 To MaxSlot Step SS
        C = 0
        Filename = Trim(File(i))                                            'Full path
        If Mode16 = True Then Filename2 = Trim(Base(i + 1))                 'Full path - 2nd file
        
        If Filename = "" Then
            If cbAllowEmpty.Value = vbUnchecked Then
                MsgBox "SLOT " & Str(i) & " is EMPTY! If you want to zero-pad these then check 'Allow Empty' option!"
                Exit Sub
            End If
        Else
            
            If Left(Filename, 1) = "%" Then
                '-- Check CMD
                If Len(Filename) < 2 Then MsgBox "Invalid CMD: " & Filename: Exit Sub
                Filename = UCase(Mid(Filename & " ", 2))
                If Left(Filename, 1) <> "F" Then MsgBox "Unknown CMD: " & Filename: Exit Sub
                If Mode16 = True And Left(Filename2, 1) <> "%" Then MsgBox "Both pairs must be %CMD in 16-bit mode!": Exit Sub
            Else
                '-- Check File
                If Exists(Filename) = False Then
                    MsgBox "Slot " & Str(i + 1) & " is unspecified or does not exist"
                    Exit Sub
                End If
                
                If (Mode16 = True) And (Exists(Filename2) = False) Then MsgBox "The 2nd file in 16-bit mode does not exit!": Exit Sub
                
                FLen = FileLen(Filename)
                
                If cbAllowShort.Value = vbUnchecked Then
                    If FLen < SlotSize Then
                        MsgBox "The file '" & Filename & "' is too short! Size: " & Str(FLen) & "!"
                        Exit Sub
                    End If
                End If
                
                If FLen > SlotSize + 2 Then
                    MsgBox "The file '" & Filename & "' is bigger than SlotSize!"
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    '--- Get a filename
    
    Filename = FileOpenSave("", 1, 2, "Add ROM"): If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    
    '--- Open the Output file
    FIO = FreeFile
    Open Filename For Output As FIO: DoEvents
    
    '===============
    ' Process Files
    '===============
    
    For i = 0 To MaxSlot Step SS
        Filename = Trim(File(i))                                                'Full path     'lblInfo.Caption = "Writing " & Filename & "...": DoEvents
         
        If Filename = "" Then
            '-- Allow empty slots. Write ZEROS to fill slot
            Padd = DefPadd: GoSub WritePadding

        Else
            If Left(Filename, 1) = "%" Then
                '-- Check CMD
                Filename = UCase(Filename) & "  "
                p = InStr(Filename, " ")
                CMD = Mid(Filename, 2, 1)
                PARM = Mid(Filename, p + 1)
                
                Select Case CMD
                    Case "F"
                        j = Val(PARM)                                           'Get value of parameter
                        Padd = Chr(j)                                           'Use decimal value after "F"
                        GoSub WritePadding                                      'Do padding
                        If Mode16 = True Then GoSub WritePadding                'Do padding again for 15-bit
                    Case Else
                        '-- unknown CMD, so just ZERO pad
                        Padd = DefPadd                                          'Use default padding value
                        GoSub WritePadding                                      'Do padding
                        If Mode16 = True Then GoSub WritePadding                'Do padding again for 15-bit
                End Select
            
            Else
                '-- Check File
                If Exists(Filename) = False Then
                    MsgBox "Slot " & Str(i + 1) & " does not exist (this should not occur)!"
                    Exit Sub
                End If
                
                Flag = False: FLen = FileLen(Filename)
                If (FLen Mod SlotSize) = 2 Then Flag = True: FLen = FLen - 2                'Includes LoadAddress!
                
                FIO2 = FreeFile
                Open Filename For Binary As FIO2
                If Flag = True Then LA = Input(2, FIO2)                                     'Read 2 bytes and discard
                
                If Mode16 = True Then
                    Filename2 = Trim(File(i + 1))                                           'Full path
                    Flag = False: FLen = FileLen(Filename2)
                    If (FLen Mod SlotSize) = 2 Then Flag = True: FLen = FLen - 2            'Includes LoadAddress!
                
                    FIO3 = FreeFile
                    Open Filename2 For Binary As FIO3
                    If Flag = True Then LA = Input(2, FIO3)                                 'Read 2 bytes and discard
                End If
                
                '--- Copy the File to Slot
                C = 0                                                                       'Zero bytes read
                
                Do While Not EOF(FIO2) And C < SlotSize
                    B = Input(1, FIO2)                                                      'Read a byte
                    C = C + 1                                                               'Count it
                    Print #FIO, B;                                                          'Write the byte
                    
                    If Mode16 = True Then                                                   'Mode16 - 2nd file
                        If Not EOF(FIO3) Then
                            B = Input(1, FIO3)                                              'Read a byte
                            Print #FIO, B;                                                  'Write the byte
                        End If
                    End If
                Loop
                
                If C < SlotSize Then
                    Padd = DefPadd                                                          'Use default padding
                    Do While C < SlotSize
                        Print #FIO, Padd;                                                   'Write padding
                        C = C + 1                                                           'Count it
                    Loop
                End If
                
                Close FIO2                                                                  'Close the file
                If Mode16 = True Then Close FIO3                                            'Mode16 - close 2nd file
                
            End If
        End If
    Next i
    
    Close FIO
    MsgBox "File successfully created!!!"
    Exit Sub

'---- Fill slot with Pad value

WritePadding:
            For j = 1 To SlotSize
                Print #FIO, Padd;
            Next j
            Return

End Sub

'===========
' FUNCTIONS
'===========

'---- Exists
' Check if a File Exists - Returns TRUE if it does
Private Function Exists(ByVal Filename As String) As Boolean
    Dim FIO As Integer
    
    On Local Error GoTo NoFile              'Open will fail if file does not exist
    FIO = FreeFile
    Open Filename For Input As FIO          'If this works then the file exists
    Close FIO
    Exists = True                           'Return TRUE
    Exit Function

NoFile:
    Close FIO
    Exists = False                          'Return FALSE
    
End Function

'---- Drag and Drop
' Processes multiple dropped files. Drops to the currently selected slot and can load
' as many files as needed up to slot 255 even if NumSlots is smaller.
'
' To enable, set OLEDropMode to "1 - Manual" for each textFN control
Private Sub txtFN_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Filename As String, N As Integer
    
    N = TopNum + Index
    'Debug.Print "Drag and Drop!"
    
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        For Each vFn In Data.Files
            Filename = (vFn)                            'vFn is name of file dropped
            File(N) = Filename                          'full path
            Base(N) = FName(Filename)                   'file without path
            FileSize(N) = 0                             'Zero file size
            FileInfo(N) = "?"                           'Mark as new entry
            
            N = N + 1                                   'Point to next slot
            If N > 255 Then Exit For                    'All slots are filled, so done
        Next vFn
    End If
    
    DrawSlots
End Sub

'---- Provide feedback to user
' If dragging a FILE then accept it, otherwise no.
Private Sub txtFN_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '0=do not allow drop, 1=inform source that data will be copied
    If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub

'---- Common File Open or Save Dialog
' You can specify a default filename, a File Filter list index (0-1), and Window Title
' MODE: 0=Open, 1=Save
' Returns a filename with full path. If cancelled will return null string
Private Function FileOpenSave(ByVal DefFile As String, ByVal Mode As Integer, FiltSet As Integer, DTitle As String) As String
    Dim Filename As String
    
    CommonDialog.CancelError = True
    On Local Error GoTo NoFile
        
    CommonDialog.DialogTitle = DTitle
    CommonDialog.Flags = cdlOFNHideReadOnly
    CommonDialog.Filename = DefFile
    
    Select Case FiltSet
        Case 0: CommonDialog.Filter = "All files (*.*)|*.*"
        Case 1: CommonDialog.Filter = "Text Files (*.TXT)|*.TXT"
        Case 2: CommonDialog.Filter = "ROM Files (*.bin, *.rom)|*.bin;*.rom"
    End Select
    
    If Mode = 0 Then CommonDialog.ShowOpen Else CommonDialog.ShowSave   'MODE: 0=Open, 1=Save
        
    If CommonDialog.Filename = "" Then Exit Function
    
    FileOpenSave = CommonDialog.Filename
    Exit Function
NoFile:

End Function

'---- Overwrite
' Checks for file and prompts to Overwrite if necessary
' Returns TRUE if file does NOT exist, or it EXISTS and user says YES.
' Returns FALSE if file EXISTS but user says NO.
Public Function Overwrite(ByVal Filename As String) As Boolean
    
    Overwrite = True    'assume ok to replace
    
    If Exists(Filename) = True Then
        If MsgBox("The file '" & Filename & "' already exists!" & Cr & "Replace it?", vbYesNo, "Overwrite File") = vbNo Then Overwrite = False
    End If
End Function

'---- FName
' Return the filename only from the end of the path
Public Function FName(ByVal Path As String) As String

Dim j As Integer

j = InStrRev(Path, "\")
If j > 0 Then
    FName = Mid(Path, j + 1)
Else
    FName = Path
End If

End Function


'---- FName
' Return the filename only from the end of the path
Public Function FNoExt(ByVal Path As String) As String

Dim j As Integer

j = InStrRev(Path, ".")
If j > 0 Then
    FNoExt = Left(Path, j - 1)           'Everything BEFORE the last PERIOD
Else
    FNoExt = Path                        'The filename has no extension (unlikely)
End If

End Function


'==== BUTTON/DROPDOWN HANDLING

Private Sub cmdNew_Click()
    NewSet
    DrawSlots
End Sub

Private Sub cboNumSlots_Click()
    SetSlotSize
End Sub

Private Sub cboTargetSize_Click()
    SetSlotSize
End Sub

Private Sub cboGroup_Click()
    SetSlotSize
End Sub

Private Sub cbAllowShort_Click()
    SetSlotSize
End Sub
Private Sub cbAllowEmpty_Click()
    SetSlotSize
End Sub
'---- Add a Binary
Private Sub cmdAdd_Click()
    AddBin
End Sub

'---- Load a Set from TXT file
Private Sub cmdLoadSet_Click()
    txtDesc.SetFocus
    LoadSet
    DrawSlots
End Sub

'---- Save a Set to TXT file
Private Sub cmdSaveSet_Click()
    SaveSet
End Sub

'---- Move selected entry DOWN
Private Sub cmdDown_Click()
    'Debug.Print "---MoveDown Btn"
    UpdateEdit
    MoveDown
    SelectN SelIndex + 1
    'Debug.Print "MoveDown Btn done."
End Sub

'---- Move selected entry UP
Private Sub cmdUp_Click()
    UpdateEdit
    MoveUp
    SelectN SelIndex - 1
End Sub

'---- Delete Selected Entry and move lower entries UP
Private Sub cmdDel_Click()
    DelSlot
    DrawSlots
End Sub

'---- Insert an Entry and move all lower entries DOWN
Private Sub cmdIns_Click()
    InsSlot
    DrawSlots
End Sub

Private Sub cmdClear_Click()
    ClearSlot SelNum
    DrawSlots
End Sub

Private Sub cmdSplit_Click()
    SplitSlot
End Sub

