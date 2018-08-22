VERSION 5.00
Object = "{20976770-692B-4564-84B5-CCC822AA2B7A}#1.4#0"; "CmdBtnX5.ocx"
Object = "{C5708EEA-B04E-45FF-9778-8A158EE64224}#1.6#0"; "Panel5.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.3#0"; "Codejock.Controls.v13.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.3#0"; "Codejock.CommandBars.v13.3.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.3#0"; "Codejock.SkinFramework.v13.3.1.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm메인
   BackColor       =   &H00B4643F&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Smart Madical Security HIS - NeoSoftBank "
   ClientHeight    =   10755
   ClientLeft      =   5235
   ClientTop       =   2460
   ClientWidth     =   14910
   Icon            =   "frm메인.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm메인.frx":C84A
   ScaleHeight     =   10755
   ScaleMode       =   0  '사용자
   ScaleWidth      =   14910
   Begin XtremeSuiteControls.Resizer Resizer1
      Height          =   11355
      Left            =   0
      TabIndex        =   3
      Top             =   -540
      Width           =   15000
      _Version        =   851971
      _ExtentX        =   26458
      _ExtentY        =   20029
      _StockProps     =   1
      BackColor       =   16315889
      Begin AriadPanelCtl.Panel pnl로그인
         Height          =   2220
         Left            =   7110
         Top             =   1590
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3916
         BorderStyle     =   5
         BorderColor     =   14396033
         BackColor       =   14396033
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin AriadPanelCtl.Panel Panel2
            Height          =   315
            Index           =   30
            Left            =   120
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "로그인ID"
            CaptionAlignment=   4
            BorderStyle     =   5
            BorderColor     =   16315889
            BackColor       =   16315889
            ForeColor       =   12282694
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AriadPanelCtl.Panel Panel2
            Height          =   315
            Index           =   0
            Left            =   120
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "사용자명"
            CaptionAlignment=   4
            BorderStyle     =   5
            BorderColor     =   16315889
            BackColor       =   16315889
            ForeColor       =   12282694
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AriadPanelCtl.Panel Panel2
            Height          =   315
            Index           =   1
            Left            =   120
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "직분"
            CaptionAlignment=   4
            BorderStyle     =   5
            BorderColor     =   16315889
            BackColor       =   16315889
            ForeColor       =   12282694
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AriadPanelCtl.Panel Panel2
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   1680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Caption         =   "진료지원실"
            CaptionAlignment=   4
            BorderStyle     =   5
            BorderColor     =   16315889
            BackColor       =   16315889
            ForeColor       =   12282694
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AriadPanelCtl.Panel Panel2
            Height          =   2235
            Index           =   3
            Left            =   3000
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   3942
            CaptionAlignment=   4
            BorderStyle     =   5
            BorderColor     =   12547143
            BackColor       =   12547143
            ForeColor       =   12282694
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin XtremeSuiteControls.PushButton btn게시판
               Height          =   405
               Left            =   120
               TabIndex        =   12
               Top             =   180
               Width           =   1335
               _Version        =   851971
               _ExtentX        =   2355
               _ExtentY        =   706
               _StockProps     =   79
               Caption         =   "게시판"
               ForeColor       =   8602147
               BackColor       =   16315889
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   4
               TextImageRelation=   1
            End
            Begin XtremeSuiteControls.PushButton btn메신져
               Height          =   405
               Left            =   120
               TabIndex        =   13
               Top             =   680
               Width           =   1335
               _Version        =   851971
               _ExtentX        =   2355
               _ExtentY        =   706
               _StockProps     =   79
               Caption         =   "메신져"
               ForeColor       =   8602147
               BackColor       =   16315889
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   4
               TextImageRelation=   1
            End
            Begin XtremeSuiteControls.PushButton btn인증서등록
               Height          =   405
               Left            =   120
               TabIndex        =   14
               Top             =   1170
               Width           =   1335
               _Version        =   851971
               _ExtentX        =   2355
               _ExtentY        =   706
               _StockProps     =   79
               Caption         =   "인증서등록"
               ForeColor       =   8602147
               BackColor       =   16315889
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   4
               TextImageRelation=   1
            End
            Begin XtremeSuiteControls.PushButton btn사용자변경
               Height          =   405
               Left            =   120
               TabIndex        =   15
               Top             =   1680
               Width           =   1335
               _Version        =   851971
               _ExtentX        =   2355
               _ExtentY        =   706
               _StockProps     =   79
               Caption         =   "사용자변경"
               ForeColor       =   8602147
               BackColor       =   16315889
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   4
               TextImageRelation=   1
            End
         End
         Begin XtremeSuiteControls.FlatEdit txt사용자ID
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   1695
            _Version        =   851971
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "FlatEdit1"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txt사용자명
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   720
            Width           =   1695
            _Version        =   851971
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "FlatEdit1"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txt직분
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   1200
            Width           =   1695
            _Version        =   851971
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "FlatEdit1"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txt진료지원실
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   1680
            Width           =   1695
            _Version        =   851971
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "FlatEdit1"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin InetCtlsObjects.Inet Inet
         Left            =   120
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.PictureBox pic해솔
         Appearance      =   0  '평면
         BackColor       =   &H00F8F5F1&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   6615
         TabIndex        =   31
         Top             =   9135
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Timer tmr긴급공지
         Interval        =   10000
         Left            =   12630
         Top             =   4155
      End
      Begin CommandButtonXCtl.CommandButtonX cmd변경아이콘
         Height          =   1440
         Left            =   3390
         TabIndex        =   18
         Top             =   -735
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":2FA95
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin VB.Timer tmr로그인
         Left            =   720
         Top             =   960
      End
      Begin FPUSpreadADO.fpSpread sprDatabase
         Height          =   2190
         Left            =   10440
         TabIndex        =   7
         Top             =   9165
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524288
         _ExtentX        =   4842
         _ExtentY        =   3863
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   1
         MaxRows         =   1
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frm메인.frx":2FB17
         StartingRowNumber=   0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txt아이디
         Height          =   480
         Left            =   7110
         TabIndex        =   1
         Top             =   1815
         Width           =   4575
         _Version        =   851971
         _ExtentX        =   8070
         _ExtentY        =   847
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "맑은 고딕"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "FlatEdit1"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txt패스워드
         Height          =   480
         Left            =   7110
         TabIndex        =   2
         Top             =   2490
         Width           =   4575
         _Version        =   851971
         _ExtentX        =   8070
         _ExtentY        =   847
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "맑은 고딕"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "FlatEdit1"
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txt데이터베이스
         Height          =   480
         Left            =   7110
         TabIndex        =   0
         Top             =   855
         Visible         =   0   'False
         Width           =   4575
         _Version        =   851971
         _ExtentX        =   8070
         _ExtentY        =   847
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "맑은 고딕"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "FlatEdit1"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin CommandButtonXCtl.CommandButtonX btnDataBase
         Height          =   480
         Left            =   10845
         TabIndex        =   4
         Top             =   1650
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   847
         DropDownPicture =   "frm메인.frx":3002B
         BorderColor     =   12282694
         BorderStyle     =   0
         Caption         =   "찾기"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12282694
         ForeColor       =   -2147483634
      End
      Begin CommandButtonXCtl.CommandButtonX cmd로그인
         Height          =   600
         Left            =   7110
         TabIndex        =   5
         Top             =   3180
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1058
         DropDownPicture =   "frm메인.frx":300AD
         FocusStyle      =   0
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8421504
      End
      Begin CommandButtonXCtl.CommandButtonX cmd인증서로그인
         Height          =   600
         Left            =   9420
         TabIndex        =   6
         Top             =   3180
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1058
         DropDownPicture =   "frm메인.frx":3012F
         FocusStyle      =   0
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8421504
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   0
         Left            =   1110
         TabIndex        =   19
         Top             =   1650
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":301B1
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   1
         Left            =   2865
         TabIndex        =   20
         Top             =   1650
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":30233
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   2
         Left            =   4620
         TabIndex        =   21
         Top             =   1650
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":302B5
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   3
         Left            =   1080
         TabIndex        =   22
         Top             =   3600
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":30337
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   4
         Left            =   2865
         TabIndex        =   23
         Top             =   3570
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":303B9
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   5
         Left            =   4620
         TabIndex        =   24
         Top             =   3570
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":3043B
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   6
         Left            =   1110
         TabIndex        =   25
         Top             =   5505
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":304BD
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   7
         Left            =   2865
         TabIndex        =   26
         Top             =   5505
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":3053F
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   8
         Left            =   4620
         TabIndex        =   27
         Top             =   5505
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":305C1
         FocusStyle      =   3
         BorderStyle     =   0
         CaptionAlignment=   5
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   9
         Left            =   1110
         TabIndex        =   28
         Top             =   7395
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":30643
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   10
         Left            =   2865
         TabIndex        =   29
         Top             =   7395
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":306C5
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin CommandButtonXCtl.CommandButtonX cmd아이콘
         Height          =   1440
         Index           =   11
         Left            =   4620
         TabIndex        =   30
         Top             =   7395
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DropDownPicture =   "frm메인.frx":30747
         FocusStyle      =   3
         BorderStyle     =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16315889
      End
      Begin VB.Image Image
         Height          =   10815
         Left            =   30
         Picture         =   "frm메인.frx":307C9
         Top             =   510
         Width           =   15000
      End
      Begin XtremeSuiteControls.Label lbl백업일자
         Height          =   180
         Left            =   7155
         TabIndex        =   32
         Top             =   10635
         Width           =   1260
         _Version        =   851971
         _ExtentX        =   2223
         _ExtentY        =   318
         _StockProps     =   79
         Caption         =   "마지막백업일자"
         ForeColor       =   5984838
         BackColor       =   11117464
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeCommandBars.ImageManager ImageManager1
         Left            =   1020
         Top             =   3120
         _Version        =   851971
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frm메인.frx":53A14
      End
      Begin XtremeSuiteControls.Label lblInfo
         Height          =   180
         Left            =   7155
         TabIndex        =   16
         Top             =   10350
         Width           =   1920
         _Version        =   851971
         _ExtentX        =   3387
         _ExtentY        =   318
         _StockProps     =   79
         Caption         =   "요양기관번호: 12345678"
         ForeColor       =   5984838
         BackColor       =   11117464
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSkinFramework.SkinFramework SkinFramework
         Left            =   14340
         Top             =   8820
         _Version        =   851971
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton3
      Height          =   1440
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   1440
      _Version        =   851971
      _ExtentX        =   2540
      _ExtentY        =   2540
      _StockProps     =   79
      Caption         =   "btn아이콘"
      BackColor       =   16777215
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextImageRelation=   0
   End
   Begin MSWinsockLib.Winsock sckLockCheck
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frm메인"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As Point) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long

' 2018-03-14 신명주
' 코어스넷 백신 기능 함수 추가_신명주
Private Declare Function Init Lib "MCaresSdk.dll" () As Boolean
Private Declare Function StartMCares Lib "MCaresSdk.dll" () As Boolean
Private Declare Function ExecuteMain Lib "MCaresSdk.dll" () As Boolean

Private lngD_메신저Pid As Long

Private ColD_Pid As New Collection
Private Type Point
    x As Long
    y As Long
    color As Long
End Type
Private Enum 아이콘구성
    데스크 = 1
    진료실 = 2
    진료지원 = 3
    청구심사 = 4
    병원관리 = 5
    문서관리 = 6
    입원수납 = 7
    병동 = 8
    외래간호 = 9
    경영관리 = 10
    원격지원 = 11
    의무기록실 = 12
    '20180521_신명주
    문서권한 = 14
    반입정보 = 15
    방역화 = 16
    비식별화 = 17
End Enum

Dim lngD_로그인횟수 As Long
Private strD_공지사항 As String
Private strD_CaptionTemp As String '16/01/07 준혁추가
Private Declare Function C_GetCertToMobile Lib "ComKSignHIS.dll" () As Long
Private Declare Function C_GenSignedData Lib "ComKSignHIS.dll" (ByVal x As String, y&) As Long
Private Declare Function C_VerifySignedDataForHIS Lib "ComKSignHIS.dll" (ByVal x As String, y&) As Long
Private Declare Function C_VerifyCert Lib "ComKSignHIS.dll" (ByVal x As String, ret&) As Long

Private Declare Function SysFreeString Lib "OleAut32.dll" (source As Any)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

Dim lngD_인증확인 As Long

Function StringFromBSTR(ByVal pointer As Long) As String
    Dim temp As String
    CopyMemory ByVal VarPtr(temp), pointer, 4
    StringFromBSTR = temp
    CopyMemory ByVal VarPtr(temp), 0&, 4
End Function
' 모바일에 저장된 인증서 요청
Private Sub sD_Login인증서인증()

Dim sPosition As Long

    Call sD_로그인이전DBupdate

    '사용자정보 가져오기



    '인증서 요청 정상 처리
    If fD_인증서요청 Then
        '사용자정보로 정상 로그인처리
        blnD_KS인증결과 = True
        If ocsLogin.ConnectServer(WinAPI_INI파일읽기("Database Server", "database_name", "C:\Sense\bin\sense.ini")) = ocsResultServerConnect0False Then
            If ocsLogin.ConnectServer(ocsWorkState.GetIniFile("Login", "DatabaseName", "")) = ocsResultServerConnect0False Then
                fG_MsgBox "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security"

                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        ocsWorkUser.mUserID = Me.txt아이디
        ResultLogin = ocsLogin.LoginProcess(Me.txt아이디, fD_사용자정보("패스워드"))

        If ResultLogin = ocsResultLogin3FalseOUT Then

            Me.MousePointer = vbDefault

            fG_MsgBox "현 사용자ID 는 퇴사 상태입니다", vbCritical + vbDefaultButton1, "SenseChart" 'Me.hwnd, "현 사용자ID 는 퇴사 상태입니다", vbCritical + vbDefaultButton1, "SenseChart"

            Me.txt아이디.SetFocus

            Exit Sub
        Else

            ocsLogin.WorkState.SetIniFile "Login", "UserID", Me.txt아이디.Text
            ocsLogin.WorkState.SetIniFile "Login", "DatabaseName", Me.txt데이터베이스.Text
            IniObjectLoad

            strD_SC프로세스 = "Express.exe"
            strD_SC타이틀 = "Smart Medical Security HIS - NeochartBank"

            Call fD_DN값    '인증서 DN값
            Call fD_전자서명생성 '전자서명
            Call fD_전자서명검증 '서명값검증
'            Call fD_인증상태검증 '유효성검증
            sD_업데이트     'DN값 저장필드 늘림

            '케이사인 DN값 비교
'            If fD_DN값 <> fD_DN값가져오기(Me.txt아이디.Text) Then
'                MsgBox "DN값이 일치하지 않습니다. 확인해주세요", vbCritical
'            End If

            '20180522_신명주 - 사용자이름과 사용자ID 값 비교
'            Dim getDN() As String
'            getDN() = Split(fD_DN값, ",")
'            getDN() = Split(getDN(0), "(")
'            getDN() = Split(getDN(0), "=")
'            If getDN(1) <> Trim(Me.txt아이디.Text) Then
'                fG_MsgBox "사용자 정보가 일치하지 않습니다.", vbCritical, vbDefaultButton1
'                Exit Sub
'            End If


            ' 20180717_신명주 병원 실인증서 DN값으로 비교되어 로그인되도록 방식 변경
            Debug.Print fD_DN값
            Debug.Print ocsWorkUser.mBigo

            If Trim(fD_DN값) <> Trim(ocsWorkUser.mBigo) Then

                fG_MsgBox "사용자 정보가 일치하지 않습니다.", vbCritical, vbDefaultButton1
                Exit Sub

            End If

            Dim cnt As Long
            cnt = ocsWorkUser.mPosition

            Select Case ocsWorkUser.mPosition
                Case 0
                    sPosition = 1
                Case 1, 7
                    sPosition = 2
                Case 9
                    sPosition = 3
                Case Else
                    sPosition = 2
            End Select

            'fD_로그인 Me.txt아이디.Text, App.Path + "\" + strD_SC프로세스

            fD_로그인 Me.txt아이디.Text, App.Path + "\" + strD_SC프로세스, sPosition

            '로그인시 화면보안 기본설정
            fD_보안설정 strD_SC프로세스, strD_SC타이틀

            sP_Loadboard
            sD_로그인완료

            ocs공용폼.MainForm = Me
        End If

        Me.MousePointer = vbDefault
    Else
        blnD_KS인증결과 = False '인증서 로그인 실패
    End If

End Sub

Private Sub cmd아이콘_MouseEnter(Index As Integer)
    Select Case cmd아이콘(Index).tag

        Case "Desk"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.데스크 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "Doctor"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료실 + 1000, 96).CreatePicture(xtpImageNormal)

                    '20150715 김석준 / Case "진료지원" -> Support로 수정
        Case "Support"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료지원 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "Chunggu"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.청구심사 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "Manage"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.병원관리 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "Documents"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서관리 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "입원수납"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.입원수납 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "병동"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.병동 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "외래간호"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.외래간호 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "통계관리"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.경영관리 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "원격지원"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.원격지원 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "의무기록실"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.의무기록실 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "문서권한"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서권한 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "반입정보"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.반입정보 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "방역화"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.방역화 + 1000, 96).CreatePicture(xtpImageNormal)

        Case "비식별화"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.비식별화 + 1000, 96).CreatePicture(xtpImageNormal)

        Case ""
            Set cmd아이콘(Index).Picture = Nothing
            cmd아이콘(Index).Visible = False

    End Select
End Sub


Private Sub cmd아이콘_MouseExit(Index As Integer)
    Select Case cmd아이콘(Index).tag

        Case "Desk"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.데스크, 96).CreatePicture(xtpImageNormal)

        Case "Doctor"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료실, 96).CreatePicture(xtpImageNormal)

                    '20150715 김석준 / Case "진료지원" -> Support로 수정
        Case "Support"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료지원, 96).CreatePicture(xtpImageNormal)

        Case "Chunggu"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.청구심사, 96).CreatePicture(xtpImageNormal)

        Case "Manage"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.병원관리, 96).CreatePicture(xtpImageNormal)

        Case "Documents"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서관리, 96).CreatePicture(xtpImageNormal)

        Case "입원수납"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.입원수납, 96).CreatePicture(xtpImageNormal)

        Case "병동"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.병동, 96).CreatePicture(xtpImageNormal)

        Case "외래간호"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.외래간호, 96).CreatePicture(xtpImageNormal)

        Case "통계관리"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.경영관리, 96).CreatePicture(xtpImageNormal)

        Case "원격지원"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.원격지원, 96).CreatePicture(xtpImageNormal)

        Case "의무기록실"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.의무기록실, 96).CreatePicture(xtpImageNormal)

        Case "문서권한"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서권한, 96).CreatePicture(xtpImageNormal)

        Case "반입정보"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.반입정보, 96).CreatePicture(xtpImageNormal)

        Case "방역화"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.방역화, 96).CreatePicture(xtpImageNormal)

        Case "비식별화"
            Set cmd아이콘(Index).Picture = ImageManager1.Icons.GetImage(아이콘구성.비식별화, 96).CreatePicture(xtpImageNormal)

        Case ""
            Set cmd아이콘(Index).Picture = Nothing
            cmd아이콘(Index).Visible = False

    End Select
End Sub


Private Sub sD_SetImage()
    Dim lngL_i As Long


    Set cmd로그인.Picture = ImageManager1.Icons.GetImage(10010, 160).CreatePicture(xtpImageNormal)
    Set cmd인증서로그인.Picture = ImageManager1.Icons.GetImage(10011, 160).CreatePicture(xtpImageNormal)

    For lngL_i = 0 To cmd아이콘.UBound

        Select Case cmd아이콘(lngL_i).tag

            Case "Desk"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.데스크, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "Doctor"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료실, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
                        '20150715 김석준 / Case "진료지원" -> Support로 수정
            Case "Support"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.진료지원, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "Chunggu"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.청구심사, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "Manage"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.병원관리, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "Documents"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서관리, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "입원수납"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.입원수납, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "병동"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.병동, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "외래간호"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.외래간호, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "통계관리"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.경영관리, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "원격지원"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.원격지원, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "의무기록실"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.의무기록실, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "문서권한"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.문서권한, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "반입정보"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.반입정보, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "방역화"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.방역화, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case "비식별화"
                Set cmd아이콘(lngL_i).Picture = ImageManager1.Icons.GetImage(아이콘구성.비식별화, 96).CreatePicture(xtpImageNormal)
                cmd아이콘(lngL_i).Visible = True
            Case ""
                Set cmd아이콘(lngL_i).Picture = Nothing
                cmd아이콘(lngL_i).Visible = False
        End Select

    Next

End Sub

Private Sub sD_setTabControl()

End Sub

Private Sub pic로그인_Click()
    sD_Login
End Sub

Private Sub btnDataBase_Click()

    If sprDatabase.Visible = False Then
        If frm암호입력.fP_암호입력("관리자 암호 입력") = "asdjkl" Then
            sprDatabase.Visible = True
        End If
    Else
        sprDatabase.Visible = False
    End If
End Sub

Private Sub btn권장사양_Click()
    If btn권장사양.Checked = False Then
        btn권장사양.Checked = True
        pic권장사양.Visible = True
    Else
        btn권장사양.Checked = False
        pic권장사양.Visible = False
    End If
End Sub

Private Sub btn사용자변경_Click()
    Dim lngL_i As Long
    Dim lngL_Pid As Variant
    Dim EndCode As Long
    Dim EndRet As Long
    Dim hprocess As Long
'    If fG_MsgBox(Me.hwnd, "로그아웃 하시겠습니까?", vbInformation + vbYesNo, "Smart Medical Security") = vbYes Then
    If fG_MsgBox("로그아웃 하시겠습니까?", vbInformation + vbYesNo, "Smart Medical Security") = vbYes Then
'        btn데이터업데이트.Visible = False
        pnl로그인.Visible = False
        sD_Clear
        cmd로그인.Enabled = True
        cmd인증서로그인.Enabled = True
        ocsDBMng.ServerDisConnect
        Me.Caption = strD_CaptionTemp ' 16/01/07 준혁 로그아웃시 요양기관명 사라지도록 추가
        '16/04/19 희빈
        'Call tab부가서비스_Click

        fD_보안해제 strD_SC프로세스, strD_SC타이틀

        '170522
        fD_로그아웃 App.Path + "\express.exe"
        'sD_UnloadMessenger

'        Call ocs센스로그.sP_InsertLogData(Log실행파일구분.메인화면, "로그아웃", WinAPI_INI파일읽기("Database Server", "database_name", "C:\Sense\bin\sense.ini") & " 로그아웃")
    End If
End Sub

Private Sub btn인증도움말_Click()

    frm인증관련도움말.Show vbModal
End Sub

Private Sub btn인증서등록_Click()
    sD_showExe "네오인증서.exe"
End Sub

Private Sub cmd로그인_Click()
    sD_Login
End Sub

Private Sub cmd인증서로그인_Click()

    sD_Login인증서인증

End Sub

Public Sub sP_Loadboard()   '16/02/11 림선 Public으로 변경.

End Sub

Private Sub sD_showExe(ByVal sProgramExeName As String)

    Dim FileName As String
    Dim FileNameHan As String
    Dim lngL_FindWindow As Long

    ChDir App.Path

    FileName = App.Path + "\" & sProgramExeName
    FileNameHan = WindowName(sProgramExeName)
    If Dir(FileName) = "" Then
        fG_MsgBox "실행할 수 없습니다.", vbCritical
        Exit Sub
    End If

    lngL_FindWindow = FindWindow("ThunderRT6FormDC", FileNameHan)
    If sProgramExeName = "Chunggu.exe" Then
        If GetPidByImage(sProgramExeName) <> 0 Then
            If lngL_FindWindow <> 0 Then
                Exit Sub
            Else
                WinAPI_KillProcess sProgramExeName
            End If
        End If
    Else
        If GetPidByImage(sProgramExeName) <> 0 Then
            If lngL_FindWindow <> 0 Then
                If MsgBox("[이미 실행중입니다 작업창을 확인하세요]" & vbNewLine & vbNewLine _
                    & "프로그램이 불안정 종료하였을경우 Ctrl+Alt+Del 키를 눌러" & vbNewLine & vbNewLine _
                    & "프로그램을 강제 종료하여주신후 다시 실행하여주세요" & vbNewLine & vbNewLine _
                    & "무시하시고 실행하려면 [취소]버튼를 클릭하세요", vbOKCancel + vbInformation) = vbOK Then
                    Exit Sub
                End If
            Else
                WinAPI_KillProcess sProgramExeName
            End If
        End If
    End If
    '실행
    Call Shell(FileName & " " & ocsDBMng.DatabaseName & "/" & ocsWorkUser.mUserID & "/" & ocsWorkUser.mUserPWD & "///" & blnD_KS인증결과, vbNormalFocus)
    ColD_Pid.Add sProgramExeName

End Sub


Private Sub cmd원격지원_Click()
        If Dir("C:\Sense\Bin\seetrol\client\SeetrolClient.exe") <> "" Then

            Shell "C:\Sense\Bin\seetrol\client\SeetrolClient.exe"
        Else

            GoWebAnyhelp

        End If
End Sub

Private Sub GoWebAnyhelp()
On Error GoTo ErrorProc:
    Dim Domain As String
    Dim s As Double

    ' 20091008 Modified by Youn -- anyhelp에서 아란타로 변경.
    'Domain = "http://happicare.anyhelp.net/"
    Domain = "medi.seetrol.com"
    'ShellExecute 0, vbNullString, Domain, vbNullString, vbNullString, 1
    'ShellExecute 0, "RemoteSupporter", Domain, vbNullString, vbNullString, 1

    Dim strIEPath As String

    strIEPath = "iexplore"


    'Shell Domain, vbNormalFocus

    'Exit Sub

    ShellExecute 0, vbNullString, strIEPath, Domain, vbNullString, 1


    Exit Sub
ErrorProc:
    MsgBox err.Description

End Sub

Private Sub cmd아이콘_Click(Index As Integer)

    If pnl로그인.Visible = False Then
        fG_MsgBox "로그인이 필요합니다.", vbInformation, "Smart Medical Security"
        Exit Sub
    End If

    Select Case cmd아이콘(Index).tag
        Case "원격지원"
            If Dir("C:\Sense\Bin\seetrol\client\SeetrolClient.exe") <> "" Then

                Shell "C:\Sense\Bin\seetrol\client\SeetrolClient.exe"

                sD_showExe "Seetrol자동설정.exe"

            Else

                GoWebAnyhelp

            End If

        Case "설정"
            ocs공용폼.Show환경설정 Me

        Case "방역화"
            Call Init
            Call StartMCares
            Call ExecuteMain

        Case "문서권한"
            Shell "C:\Windows\softcamp\SHIELDEX\SDUserInfo.exe"
'            fD_권한정보호출
        Case "반입정보"
            Shell "C:\Windows\softcamp\SHIELDEX\SDAgentUI.exe"
'            fD_반입내역

        Case "비식별화"
            frm비식별화.sP_Show "http://the-51.synology.me/main/Login.do"

        Case Else
            '16/03/23 한샘 설정된 실행파일 아이콘이 없는 버튼 클릭 시 실행되지 않도록 조건 추가
            If cmd아이콘(Index).tag = "" Then Exit Sub

'            Call ocs센스로그.sP_InsertLogData(Log실행파일구분.메인화면, "프로그램실행", cmd아이콘(Index).tag & " 프로그램 실행")

            sD_showExe cmd아이콘(Index).tag & ".exe"

    End Select

End Sub


Private Sub Form_Activate()
    txt패스워드.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strL_담당자 As String
    If KeyCode = vbKeyF5 Then   '화면보안
        fD_보안설정 strD_SC프로세스, strD_SC타이틀
        Exit Sub
    ElseIf KeyCode = vbKeyF6 Then   '보안해제
        fD_보안해제 strD_SC프로세스, strD_SC타이틀
        Exit Sub
    ElseIf KeyCode = vbKeyF11 Then
        If pnl로그인.Visible = False Then Exit Sub
'        frmNew오픈프로세스.Show
    ElseIf KeyCode = vbKeyF12 Then
        If pnl로그인.Visible = False Then Exit Sub

        strL_담당자 = fG_InputBox("담당자 이름을 입력해주세요.(필수입력)")

        If LenH(Left(strL_담당자, 1)) = 2 And Len(strL_담당자) <= 3 Then

            frm방문프로세스.Show
            frm방문프로세스.strD_담당자 = strL_담당자

        Else

            fG_MsgBox "담당자 이름 설정이 잘못되었습니다.", vbInformation

        End If
    End If
End Sub


Private Sub Form_Load()
    lbl백업일자.Caption = ""
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    SkinFramework.LoadSkin App.Path & "\SenseStyle.cjstyles", ""
    SkinFramework.ApplyWindow Me.hwnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    sG_SetControlsFont Me.Controls, MainFont, Me.hwnd
    변경아이콘 = ""

    '15/04/09 재민추가

    Call sD_INI체크

    Call sD_setTabControl
    Call sD_IconTagSetting
    Call sD_Clear
    Call sP_Loadboard   '16/02/12 림선 명칭 sD -> sP로 변경.
    tmr로그인.Enabled = False
'    btn데이터업데이트.Visible = False
'    sD_SetVersion
    strD_CaptionTemp = Me.Caption '16/01/07 준혁 로그아웃 시 캡션초기화를위해 추가
'    frm홈페이지.Show
End Sub

Private Sub sD_INI체크()
On Error GoTo err:
    Dim strL_ini경로 As String

    If isDebugMode = True Then Exit Sub

    strL_ini경로 = WinAPI_INI파일읽기("자동업데이트경로", "INI경로", "C:\sense\bin\sense.ini", "")

    If Dir(strL_ini경로 & "\sense.ini") = "" Then Exit Sub

    If WinAPI_INI파일읽기("Database Server", "server_name", strL_ini경로 & "\sense.ini", "") <> "" Then
        Call WinAPI_INI파일쓰기("Database Server", "server_name", WinAPI_INI파일읽기("Database Server", "server_name", strL_ini경로 & "\sense.ini", ""), "C:\sense\bin\sense.ini")
    End If

    If WinAPI_INI파일읽기("Database Server", "server_id", strL_ini경로 & "\sense.ini", "") <> "" Then
        Call WinAPI_INI파일쓰기("Database Server", "server_id", WinAPI_INI파일읽기("Database Server", "server_id", strL_ini경로 & "\sense.ini", ""), "C:\sense\bin\sense.ini")
    End If

    If WinAPI_INI파일읽기("Database Server", "server_pwd", strL_ini경로 & "\sense.ini", "") <> "" Then
        Call WinAPI_INI파일쓰기("Database Server", "server_pwd", WinAPI_INI파일읽기("Database Server", "server_pwd", strL_ini경로 & "\sense.ini", ""), "C:\sense\bin\sense.ini")
    End If

    If WinAPI_INI파일읽기("Database Server", "database_name", strL_ini경로 & "\sense.ini", "") <> "" Then
        Call WinAPI_INI파일쓰기("Database Server", "database_name", WinAPI_INI파일읽기("Database Server", "database_name", strL_ini경로 & "\sense.ini", ""), "C:\sense\bin\sense.ini")
    End If

    Exit Sub
err:

End Sub

Private Sub sD_SaveIconTagSetting()
    Dim strL_Tag As String
    Dim lngL_i As Long

    For lngL_i = 0 To cmd아이콘.UBound
        If lngL_i = 0 Then
            strL_Tag = cmd아이콘(lngL_i).tag
        Else
            strL_Tag = strL_Tag & "|" & cmd아이콘(lngL_i).tag
        End If
    Next

End Sub

Private Sub sD_IconTagSetting(Optional 권한적용 As Boolean = False)
    Dim strL_아이콘배열() As String
    Dim lngL_i As Long
    Dim 아이콘miss As Boolean
    Dim strL_아이콘순서 As String   ' 20150825 조영환 처리내용 : 디폴트 변수 추가

    ' 20150825 조영환 처리내용 : 사용자별로 병원관리 권한설정에서 지정한 아이콘만 보이도록 수정
'    strL_아이콘배열 = Split(ocsWorkState.GetIniFile("MainICon", "순서", "Desk|Doctor|Support|Chunggu|Manage|Documents|입원수납|병동|의무기록실|통계관리|원격지원|설정"), "|")
    If INIRead(App.Path & "\Sense.ini", "메인화면실행", "입원사용여부", 99) = 0 Then
        'strL_아이콘순서 = "Desk|Doctor|Support|Chunggu|Manage|Documents|통계관리|원격지원|의무기록실"
        strL_아이콘순서 = "Desk|Doctor|Support|Manage|Documents"
    Else
'        strL_아이콘순서 = "Desk|Doctor|Support|Chunggu|Manage|Documents|입원수납|병동|외래간호|통계관리|원격지원|의무기록실"
        strL_아이콘순서 = "Desk|Doctor|Support|Manage|Documents|통계관리|방역화|문서권한|반입정보||비식별화"
    End If

    If Right(strL_아이콘순서, 2) = "||" Then
        strL_아이콘순서 = Left(strL_아이콘순서, Len(strL_아이콘순서) - 2)
    End If

    If Right(strL_아이콘순서, 1) = "|" Then
        strL_아이콘순서 = Left(strL_아이콘순서, Len(strL_아이콘순서) - 1)
    End If

    strL_아이콘배열 = Split(strL_아이콘순서, "|")

    '16/12/29 준혁 아이콘tag 초기화 추가
    For lngL_i = 0 To cmd아이콘.UBound

        cmd아이콘(lngL_i).tag = ""

    Next

    For lngL_i = 0 To UBound(strL_아이콘배열)

        cmd아이콘(lngL_i).tag = strL_아이콘배열(lngL_i)
    Next

    Call sD_SetImage

End Sub

Private Sub sD_Clear()
    txt데이터베이스.Text = WinAPI_INI파일읽기("Database Server", "database_name", "C:\Sense\bin\sense.ini")
    txt아이디.Text = ocsWorkState.GetIniFile("Login", "UserID", "")
    txt패스워드.Text = ""
    lblInfo.Caption = ""
End Sub

Private Sub sD_setSpread()
    Dim rst As ADODB.Recordset
    Dim rstL_Table As ADODB.Recordset
    Dim colL_MediData As Collection
    Dim lngL_i As Long
    sprDatabase.Visible = False

    If ocsDataList.IsServerConnect = False Then Exit Sub

    SQL_STR = "exec sp_databases "
    Set rst = ocsDataList.GetRecordSet(SQL_STR)
    SetSprCol sprDatabase, 1, "데이터베이스", 23, TypeHAlignCenter, , , CellTypeStaticText

    Do Until rst.EOF
        lngL_i = lngL_i + 1
        sprDatabase.Row = lngL_i
        sprDatabase.MaxRows = sprDatabase.Row
        sprDatabase.Text = rst.Fields("DATABASE_NAME")
        rst.MoveNext
    Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo err:


    '170522
    fD_로그아웃 App.Path + "\express.exe"
err:
    End
End Sub

Private Sub sD_Login()
    Dim ResultLogin As ocsEnumResultLogin

    MousePointer = vbHourglass

    ResultLogin = ocsResultLogin1FalseID

    UnloadMe = False

    ocsLogin.blnD_NEO로그인상태 = 1

    If txt데이터베이스.Visible = True Then
        If ocsLogin.ConnectServer(Me.txt데이터베이스.Text) = ocsResultServerConnect0False Then

            fG_MsgBox "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security"

            Me.MousePointer = vbDefault

            Exit Sub
        Else

           ' Call sD_로그인이전DBupdate

        End If
    Else
        If ocsLogin.ConnectServer(WinAPI_INI파일읽기("Database Server", "database_name", "C:\Sense\bin\sense.ini")) = ocsResultServerConnect0False Then
            If ocsLogin.ConnectServer(ocsWorkState.GetIniFile("Login", "DatabaseName", "")) = ocsResultServerConnect0False Then
                fG_MsgBox "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "서버 연결 실패.", vbCritical + vbDefaultButton1, "Smart Medical Security"

                Me.MousePointer = vbDefault

                Exit Sub
            Else
                'Call sD_로그인이전DBupdate
            End If
        Else

            'Call sD_로그인이전DBupdate

        End If
    End If

    Call sD_로그인이전DBupdate

    blnD_KS인증결과 = False

    ResultLogin = ocsLogin.LoginProcess(Me.txt아이디, Me.txt패스워드)

    If ResultLogin = ocsResultLogin1FalseID Then

        Me.MousePointer = vbDefault

        fG_MsgBox "사용자ID 가 잘못되었습니다", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "사용자ID 가 잘못되었습니다", vbCritical + vbDefaultButton1, "Smart Medical Security"

        Me.txt아이디.SetFocus

        Exit Sub

    ElseIf ResultLogin = ocsResultLogin2FalsePWD Then
        Me.MousePointer = vbDefault

        fG_MsgBox "패스워드가 잘못되었습니다.", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "패스워드가 잘못되었습니다.", vbCritical + vbDefaultButton1, "Smart Medical Security"

        Me.txt패스워드.SetFocus

        Exit Sub

    ElseIf ResultLogin = ocsResultLogin3FalseOUT Then

        Me.MousePointer = vbDefault

        fG_MsgBox "현 사용자ID 는 퇴사 상태입니다", vbCritical + vbDefaultButton1, "Smart Medical Security" 'Me.hwnd, "현 사용자ID 는 퇴사 상태입니다", vbCritical + vbDefaultButton1, "Smart Medical Security"

        Me.txt아이디.SetFocus

        Exit Sub

    Else

        ocsLogin.WorkState.SetIniFile "Login", "UserID", Me.txt아이디.Text
        ocsLogin.WorkState.SetIniFile "Login", "DatabaseName", ocsDBMng.DatabaseName
        IniObjectLoad
        If ocsWorkUser.mPosition <> ocsUserPosition7관리자 Then
            Call sD_Get인증서      '15/10/21 림선 주석 처리
        End If
        If ocsWorkState.GetIniFile("메인화면실행", "입원사용여부", 99) = 99 Then
            ocsWorkState.SetIniFile "메인화면실행", "입원사용여부", m입원여부
            If m입원여부 = 0 Then
                fG_MsgBox "프로그램을 재실행 해주세요.", vbInformation
            End If
        End If

        Call sD_Chk사용기간

        If fD_활성화여부 = False Then End

        sP_Loadboard    '16/02/12 림선 명칭 sD -> sP로 변경.

        sD_로그인완료

        '2015-06-30 재민
        If ocsWorkUser.mPosition <> ocsUserPosition7관리자 And UCase(Trim(Me.txt패스워드)) <> "ASDJKLJKL" Then

        End If

        ocs공용폼.MainForm = Me

    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub sD_로그인완료()
    Dim rst As ADODB.Recordset

    sD_IconTagSetting True

    lblInfo.Caption = "요양기관기호: " & ocsWorkState.Hospital.mHospitalCode
    lblInfo.Caption = lblInfo.Caption & " 요양기관명칭: " & ocsWorkState.Hospital.mHospitalName

    txt사용자ID = ocsWorkUser.mUserID
    txt사용자명 = ocsWorkUser.mUserName
    txt직분 = fD_사용자정보("직분")
    txt진료지원실 = fD_사용자정보("진료지원실")

    pnl로그인.Height = 0
    pnl로그인.Width = 0
    pnl로그인.Visible = True
    tmr로그인.Interval = 1
    tmr로그인.Enabled = True

    cmd로그인.Enabled = False
    cmd인증서로그인.Enabled = False

    lngD_로그인횟수 = lngD_로그인횟수 + 1

End Sub

Private Sub tmr로그인_Timer()

    If pnl로그인.Height < 2215 Then
        pnl로그인.Height = pnl로그인.Height + 60
    End If

    If pnl로그인.Width < 4575 Then
        pnl로그인.Width = pnl로그인.Width + 120
    End If

    If pnl로그인.Width >= 4575 And pnl로그인.Height >= 2215 Then
        tmr로그인.Enabled = False
    End If

End Sub

Private Function fD_사용자정보(ByVal 검색구분 As String) As String
    Dim rst As ADODB.Recordset

    fD_사용자정보 = ""

    Select Case 검색구분
        Case "직분"
            SQL_STR = "select 명칭 as 사용자정보 from tb_코드정보 where 코드종류 = '11' and 코드 = '" & ocsWorkUser.mPosition & "' "
        Case "진료지원실"
            SQL_STR = "select 진료지원실 as 사용자정보 from tb_진료지원 where 진료지원코드 = '" & ocsWorkUser.mJinSupportCode & "' "
        Case "패스워드"
            SQL_STR = "select 사용자PWD as 사용자정보 from tb_사용자정보 where 사용자ID = '" & ocsWorkUser.mUserID & "' "
        Case "라이센스"
            SQL_STR = "select 면허번호 as 사용자정보 from tb_사용자정보 where 사용자ID = '" & ocsWorkUser.mUserID & "' "
    End Select

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount <> 0 Then
        fD_사용자정보 = Trim(rst.Fields("사용자정보"))
    End If

    Set rst = Nothing

End Function

Private Sub txt패스워드_GotFocus()
    If txt패스워드.Text = "패스워드를 입력해주세요." Then
        txt패스워드.Text = ""
        txt패스워드.PasswordChar = "*"
    End If
End Sub

Private Sub txt패스워드_LostFocus()
    If txt패스워드.Text = "" Then
        txt패스워드.Text = "패스워드를 입력해주세요."
        txt패스워드.PasswordChar = ""
    End If
End Sub

Private Sub txt아이디_GotFocus()
    If txt아이디.Text = "아이디를 입력해주세요." Then
        txt아이디.Text = ""
    End If
End Sub

Private Sub txt아이디_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        txt패스워드.SetFocus
    End If

End Sub

Private Sub txt아이디_LostFocus()
    If txt아이디.Text = "" Then
        txt아이디.Text = "아이디를 입력해주세요."
    End If
End Sub

Private Sub txt패스워드_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmd로그인.Value = True
    End If
End Sub

'15/05/06 정수 FindWindow
Private Function WindowName(ProgramName As String)
    If ProgramName = "Desk.exe" Then WindowName = "데스크 - Smart Medical Security (주)NeoSoftBank"
    If ProgramName = "Doctor.exe" Then WindowName = "진료실 - Smart Medical Security (주)NeoSoftBank"
    If ProgramName = "Support.exe" Then WindowName = "진료지원 - Smart Medical Security (주)NeoSoftBank"
    'If ProgramName = "Chunggu.exe" Then WindowName = "SENSE CHART - 청구심사 (주)NeosoftBank"
    If ProgramName = "Manage.exe" Then WindowName = "병원관리 - Smart Medical Security (주)NeoSoftBank"
    If ProgramName = "Documents.exe" Then WindowName = "Smart Medical Security - 문서관리 (주)NeosoftBank"
End Function

Private Sub sD_로그인이전DBupdate()
On Error GoTo ErrorProc:
    Dim strL_ErrMsg As String

    SQL_STR = " " & vbNewLine
    SQL_STR = SQL_STR & " IF NOT EXISTS(SELECT NAME FROM SYSOBJECTS WHERE NAME='TB_전자인증') BEGIN" & vbNewLine
    SQL_STR = SQL_STR & "     CREATE TABLE [dbo].[TB_전자인증]( " & vbNewLine
    SQL_STR = SQL_STR & "         [인덱스] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [챠트번호] [nvarchar](10) NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [챠트인덱스] [nvarchar](10) NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [입원ID] [int] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [진료일자] [nvarchar](10) NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [전자서명일자] [datetime] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [서명용인증서] [ntext] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [전자서명값] [ntext] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [인증서분류] [int] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [원문데이터] [ntext] NULL, " & vbNewLine
    SQL_STR = SQL_STR & "         [전자서명자] [nvarchar](12) NULL, " & vbNewLine
    SQL_STR = SQL_STR & "      CONSTRAINT [PK_전자인증] PRIMARY KEY CLUSTERED  " & vbNewLine
    SQL_STR = SQL_STR & "     ( " & vbNewLine
    SQL_STR = SQL_STR & "         [인덱스] ASC " & vbNewLine
    SQL_STR = SQL_STR & "     )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & vbNewLine
    SQL_STR = SQL_STR & "     ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] " & vbNewLine
    SQL_STR = SQL_STR & " END  " & vbNewLine
    If ocsDBMng.RunSQL(SQL_STR) <> 1 Then
        strL_ErrMsg = "TB_전자인증"
        GoTo ErrorProc
    End If

   Exit Sub
ErrorProc:
    If strL_ErrMsg <> "" Then
        ErrorDesc = "[mdlDatabaseUpdate/sG_DatabaseAddTable/" & strL_ErrMsg & "] " & err.Number & "-" & err.source & "/" & err.Description
        Debug.Print ErrorDesc
        ErrorWriteLog ErrorDesc
    End If

    Resume Next
End Sub

Private Sub sD_ProgramUpdate()
On Error GoTo err:
    Dim Server_Path As String
    Dim Server_ver As String
    Dim Local_Path As String
    Dim Local_FilePath As String
    Dim Local_ver As String
    Dim utdShellOpStruct As SHFILEOPSTRUCT
    Dim rst As ADODB.Recordset
    Dim UserName As String
    Dim UserPassword As String
    Dim runLine As String

    Open App.Path & "\version.txt" For Input As #1
        Line Input #1, Local_ver
    Close #1

    If Server_ver = "" Or Local_ver = "" Then Exit Sub

    If Not Left(Server_ver, Len(Server_ver) - 2) <> Left(Local_ver, Len(Local_ver) - 2) Then Exit Sub

    If IsUserAnAdmin = 0 Then
        '16/05/11 한샘 우클릭하여 관리자권한으로 실행하라는 메세지 문구 추가
        fG_MsgBox "업데이트 진행을 위해 센스 실행시 실행파일에서 우클릭하여 관리자 권한으로 실행해주세요.(프로그램이 종료됩니다.)", vbInformation
        End
    End If

    fG_MsgBox "프로그램이 업데이트 되었습니다.[필수업데이트]" & vbNewLine & vbNewLine _
                     & "[확인]을 클릭후 다운로드 완료후에 [다음]을 클릭하여 업데이트를 진행하세요.", vbInformation, "업데이트"

    If Dir(Local_FilePath, vbDirectory) = "" Then MkDir Local_FilePath
    Local_FilePath = Local_FilePath & "\" & Server_ver & ".exe"
    If Dir(Local_FilePath) <> "" Then Kill Local_FilePath

    With utdShellOpStruct

        .hwnd = Me.hwnd
        .wfunc = FO_COPY
        .pfrom = Server_Path & "\" & "SenseUpdate.exe"
        .pto = Local_FilePath

    End With

    SHFileOperation utdShellOpStruct

    If Dir(Local_FilePath) <> "" Then
        Shell Local_FilePath
        End
    End If

    End

    Exit Sub
err:

End Sub

Private Sub sD_Get인증서()

    Dim rst As ADODB.Recordset
    Dim FileRst As ADODB.Recordset
    Dim rstStream As ADODB.Stream
    Dim strL_경로 As String
    Dim splitdata() As String
    Dim lngL_i, lngL_Split As Long
    Dim strL_확인경로 As String

    Set FileRst = New ADODB.Recordset

    SQL_STR = "SELECT NAME FROM SYSOBJECTS WHERE NAME='TB_네오인증서'"

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount = 0 Then Exit Sub

    SQL_STR = "Select 일련번호 from tb_네오인증서 "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount = 0 Then Exit Sub

    SQL_STR = "Select distinct 폴더명 from tb_네오인증서 "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount = 0 Then Exit Sub

    strL_경로 = 인증서Path & rst.Fields("폴더명")

    splitdata = Split(strL_경로, "\")

    For lngL_i = 0 To UBound(splitdata)

        strL_확인경로 = ""
        For lngL_Split = 0 To lngL_i
            If lngL_Split = lngL_i Then
                strL_확인경로 = strL_확인경로 & splitdata(lngL_Split)
            Else
                strL_확인경로 = strL_확인경로 & splitdata(lngL_Split) & "\"
            End If
        Next

        '16/05/02 준혁 AppData 폴더가 '숨김' 인 경우가 있어 vbHidden 추가
        If Dir(strL_확인경로, vbDirectory + vbHidden) = "" Then
            MkDir strL_확인경로
        End If

    Next


    SQL_STR = "Select 인증서, 파일명 from tb_네오인증서 "

    FileRst.Open SQL_STR, ocsDBMng.sConnect, adOpenForwardOnly, adLockReadOnly

    Set FileRst.ActiveConnection = Nothing

    Do Until FileRst.EOF

        Set rstStream = New ADODB.Stream
        rstStream.Type = adTypeBinary
        rstStream.Open
        rstStream.Write FileRst("인증서")
        rstStream.SaveToFile 인증서Path & rst.Fields("폴더명") & "\" & FileRst.Fields("파일명").Value, adSaveCreateOverWrite
        If Dir("C:\Program Files", vbDirectory) <> "" Then
            If Dir("C:\Program Files\NPKI\", vbDirectory) = "" Then
                MkDir "C:\Program Files\NPKI"
            End If
            If Dir("C:\Program Files\NPKI\KICA\", vbDirectory) = "" Then
                MkDir "C:\Program Files\NPKI\KICA"
            End If
            If Dir("C:\Program Files\NPKI\KICA\USER\", vbDirectory) = "" Then
                MkDir "C:\Program Files\NPKI\KICA\USER"
            End If
            If Dir("C:\Program Files\NPKI\KICA\USER\" & rst.Fields("폴더명"), vbDirectory) = "" Then
                MkDir "C:\Program Files\NPKI\KICA\USER\" & rst.Fields("폴더명")
            End If
            rstStream.SaveToFile "C:\Program Files\NPKI\KICA\USER\" & rst.Fields("폴더명") & "\" & FileRst.Fields("파일명").Value, adSaveCreateOverWrite
        End If

        rstStream.Close
        Set rstStream = Nothing

        FileRst.MoveNext
    Loop

End Sub

Private Function 인증서Path() As String

    Dim UserName As String
    Dim Buffer As String * 25

    Dim dl, ariesong
    Dim s As String
    aries = 200
    s = Space$(aries)
    dl = Len(s)
    Call GetWindowsDirectory(s, dl)
    tp = Left(Trim(s), 3)

    ret = GetUserName(Buffer, 25)
    UserName = Left$(Buffer, InStr(Buffer, Chr(0)) - 1)

    tpath = "C:\Users\" + Trim(UserName) + "\appData\LocalLow\NPKI\KiCA\USER\"

    인증서Path = tpath

End Function

Private Sub sD_DeleteLogData()
    Dim FileName

    If Right(Format(Date, "YYYY-MM-DD"), 2) <> "01" Then Exit Sub

    If ocsWorkState.GetIniFile("Log", "로그삭제일자", "0") = Format(Date, "YYYY-MM-DD") Then Exit Sub

    FileName = Dir(App.Path & "\..\Resource\Log\", vbNormal)

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." And FileName <> "" Then
            Kill App.Path & "\..\Resource\Log\" & FileName
        End If
        FileName = Dir

    Loop

    FileName = Dir(App.Path & "\..\Resource\Error\", vbNormal)

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." And FileName <> "" Then
            Kill App.Path & "\..\Resource\Error\" & FileName
        End If
        FileName = Dir

    Loop

    FileName = Dir(App.Path & "\..\Resource\Log\", vbNormal)

    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." And FileName <> "" Then
            Kill App.Path & "\..\Resource\Log\" & FileName
        End If
        FileName = Dir

    Loop

    ocsWorkState.SetIniFile "Log", "로그삭제일자", Format(Date, "YYYY-MM-DD")

End Sub

Private Sub sD_로그축소()
    Dim rst As ADODB.Recordset

    If ocs연결여부 = False Then Exit Sub

    SQL_STR = "Select 로그축소일자 from ocs_server.medichart.dbo.SS_기관정보 where 요양기관코드 = '" & ocsWorkState.Hospital.mHospitalCode & "' "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.State = 0 Then Exit Sub

    If rst.RecordCount = 0 Then Exit Sub

    If DateAdd("m", 1, rst.Fields("로그축소일자")) <= Format(Date, "YYYY-MM-DD") Or rst.Fields("로그축소일자") & "" = "" Then

        SQL_STR = "ALTER DATABASE [" & ocsDBMng.DatabaseName & "] SET RECOVERY SIMPLE WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "DBCC SHRINKDATABASE([" & ocsDBMng.DatabaseName & "], 10, TRUNCATEONLY)"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "ALTER DATABASE [" & ocsDBMng.DatabaseName & "] SET RECOVERY FULL WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR

        SQL_STR = "ALTER DATABASE [SenseLogData] SET RECOVERY SIMPLE WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "DBCC SHRINKDATABASE([SenseLogData], 10, TRUNCATEONLY)"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "ALTER DATABASE [SenseLogData] SET RECOVERY FULL WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR

        SQL_STR = "ALTER DATABASE [SenseReceiptHistory] SET RECOVERY SIMPLE WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "DBCC SHRINKDATABASE([SenseReceiptHistory], 10, TRUNCATEONLY)"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "ALTER DATABASE [SenseReceiptHistory] SET RECOVERY FULL WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR

        SQL_STR = "ALTER DATABASE [SENSEIMAGE] SET RECOVERY SIMPLE WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "DBCC SHRINKDATABASE([SENSEIMAGE], 10, TRUNCATEONLY)"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "ALTER DATABASE [SENSEIMAGE] SET RECOVERY FULL WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR

        SQL_STR = "ALTER DATABASE [SenseImage" & Format(Date, "YYYY") & "] SET RECOVERY SIMPLE WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "DBCC SHRINKDATABASE([SenseImage" & Format(Date, "YYYY") & "], 10, TRUNCATEONLY)"
        ocsDBMng.RunSQL SQL_STR
        SQL_STR = "ALTER DATABASE [SenseImage" & Format(Date, "YYYY") & "] SET RECOVERY FULL WITH NO_WAIT"
        ocsDBMng.RunSQL SQL_STR

        SQL_STR = "Update ocs_server.medichart.dbo.SS_기관정보 Set 로그축소일자 = '" & Format(Date, "YYYY-MM-DD") & "' where 요양기관코드 = '" & ocsWorkState.Hospital.mHospitalCode & "' "

        ocsDBMng.RunSQL SQL_STR

    End If

End Sub

Private Sub sD_ChkHospitalLock()

    Dim intL_Count As Integer

    LOCK_INS.intLOCK = -1
    LOCK_PRO.intLOCK = -1
    LOCK_IND.intLOCK = -1
    LOCK_TA.intLOCK = -1

    sD_CheckChungLock PartInsurance

End Sub

Private Sub sD_CheckChungLock(ByVal intV_Part As Integer)
On Error GoTo err:
    sckLockCheck.tag = intV_Part
    sckLockCheck.Connect "www.neochart.co.kr", 80
err:
End Sub

Private Sub sckLockCheck_Close()
Dim strL_Msg As String
    sckLockCheck.Close
    Select Case Val(sckLockCheck.tag)
        Case PartInsurance
            sD_CheckChungLock PartProtected
        Case PartProtected
            sD_CheckChungLock PartAutomobile
        Case PartAutomobile
            sD_CheckChungLock PartIndustrial
        Case Else
    End Select
End Sub

Private Sub sckLockCheck_DataArrival(ByVal bytesTotal As Long)

    Dim strL_Data As String
    Dim strL_Record() As String

    sckLockCheck.GetData strL_Data, vbString, bytesTotal
    sckLockCheck.Close
    If strL_Data <> "20020" Then
        MsgBox "ERROR : 0EX15a57WE 예기치않은 치명적인 오류가 발생하였습니다.", vbCritical
        End
    End If

End Sub

Private Function fD_테스트서버() As Boolean
    Dim rst As ADODB.Recordset

    fD_테스트서버 = False

    If ocs연결여부 = False Then Exit Function

    SQL_STR = " select 테스트서버 from ocs_server.medichart.dbo.SS_기관정보 where 요양기관코드 = '" & ocsWorkState.Hospital.mHospitalCode & "' "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.State = 0 Then Exit Function

    If rst.RecordCount = 0 Then Exit Function

    If rst.Fields("테스트서버") = "0" Then
        fD_테스트서버 = False
    Else
        fD_테스트서버 = True
    End If
End Function

Private Sub sD_Chk사용기간()
    Dim rst As ADODB.Recordset

    If ocs연결여부 = False Then Exit Sub

    SQL_STR = " select 요양기관코드 from ocs_server.medichart.dbo.SS_기관정보 where 요양기관코드 = '" & ocsWorkState.Hospital.mHospitalCode & "' and 테스트서버 = 1 and isnull(사용기간, '9999-99-99') <= '" & Format(ocsDBMng.GetSystemDate, "YYYY-MM-DD") & "' "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount <> 0 Then End

End Sub

Private Function fD_활성화여부() As Boolean
    Dim rst As ADODB.Recordset

    fD_활성화여부 = False

    If ocs연결여부 = False Then
        fD_활성화여부 = True
        Exit Function
    End If

    SQL_STR = " select 삭제 from ocs_server.medichart.dbo.SS_기관정보 where 요양기관코드 = '" & ocsWorkState.Hospital.mHospitalCode & "' "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.State = 0 Then Exit Function

    If rst.RecordCount = 0 Then Exit Function

    If rst.Fields("삭제") = "1" Then
        fD_활성화여부 = False
    Else
        fD_활성화여부 = True
    End If
End Function
'20160620 DN값 읽어오기
Private Function fD_DN값가져오기(strR_아이디 As String) As String

    Dim rst As ADODB.Recordset

    fD_DN값가져오기 = ""

    SQL_STR = "SELECT 영문이름 FROM TB_사용자정보 "
    SQL_STR = SQL_STR & "WHERE 사용자ID = '" & strR_아이디 & "'"

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.RecordCount <> 0 Then
        strD_KSDN값 = rst.Fields("영문이름")
        If strD_KSDN값 <> "" Then fD_DN값가져오기 = strD_KSDN값
        fD_DN값가져오기 = strD_KSDN값
    Else
       Exit Function
    End If

End Function
'20160620 DN값 읽어오기
Private Function sD_업데이트() As String

    Dim rst As ADODB.Recordset


    SQL_STR = " alter table TB_사용자정보 ALTER COLUMN 영문이름 nvarchar(100) "

    Set rst = ocsDBMng.GetRecordSet(SQL_STR)

    If rst.State = 0 Then Exit Function

End Function
