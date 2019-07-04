VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "RFIDUSB Tool"
   ClientHeight    =   8376
   ClientLeft      =   180
   ClientTop       =   -2736
   ClientWidth     =   15240
   DrawMode        =   1  '黑色
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8376
   ScaleWidth      =   15240
   Begin VB.CommandButton btn_Close 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btn_Timer1_Dis 
      Caption         =   "Timer1 Dis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton btn_Timer1_En 
      Caption         =   "Timer En"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   360
   End
   Begin VB.Frame frameIntrup_RW 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   15045
      Begin VB.TextBox TextIR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '兩者皆有
         TabIndex        =   33
         Top             =   240
         Width           =   12255
      End
      Begin VB.CommandButton Iread 
         Caption         =   "Tx/RX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Tx / Rx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox TextPID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Text            =   "C02E"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TextVID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Text            =   "1325"
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame FramTagRW 
      Caption         =   "RFID Tag W/R operation (Hex 00~FF Only)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6372
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   15015
      Begin VB.TextBox txt_User_New 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txt_EPC_Data 
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   56
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox txt_Reserved_KillPW_New 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   52
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt_Reserved_AccessPW_New 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   51
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt_User 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   47
         ToolTipText     =   "Only HEX format  accepted"
         Top             =   3360
         Width           =   3375
      End
      Begin VB.ListBox lstResults 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   816
         Left            =   2040
         TabIndex        =   46
         Top             =   5400
         Width           =   6855
      End
      Begin VB.TextBox txt_TID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   45
         Top             =   2880
         Width           =   6855
      End
      Begin VB.TextBox txt_Reserved_AccessPW 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   43
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt_Reserved_KillPW 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   42
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame frameScan 
         Caption         =   "Scan / StopScan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   9720
         TabIndex        =   27
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton btn_RF_init 
            Caption         =   "RF Init"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton btn_Clear 
            Caption         =   "Clear Log"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   32
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton btn_ScanStop 
            Caption         =   "StopScan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   31
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton btn_ScanStart 
            Caption         =   "Scan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton btn_Test 
            Caption         =   "Test"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton contion 
            Caption         =   "Connect"
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txt_EPC_Data_New 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         TabIndex        =   23
         Top             =   2400
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Module Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2772
         Left            =   9720
         TabIndex        =   12
         Top             =   1920
         Width           =   4575
         Begin VB.CommandButton BTN_M_WriteUser 
            Caption         =   "Write User"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   22
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_WriteEPC 
            Caption         =   "Write EPC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   21
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_WriteReservedAccess 
            Caption         =   "W Access PW"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   20
            ToolTipText     =   "Write Reserved Access"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_ReadUser 
            Caption         =   "Read User"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_ReadEPC 
            Caption         =   "Read EPC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_ReadTID 
            Caption         =   "Read TID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_ReadReserved 
            Caption         =   "Read Reserved"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_Read 
            Caption         =   "Read Tag"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_Write 
            Caption         =   "Write Tag"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton BTN_M_WriteReservedKill 
            Caption         =   "Write Kill PW"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   13
            ToolTipText     =   "Write Reserved Kill"
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Module Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1452
         Left            =   9720
         TabIndex        =   9
         Top             =   4680
         Width           =   4575
         Begin VB.HScrollBar HS_RxSensetivity 
            Height          =   375
            Left            =   2160
            Max             =   -51
            Min             =   -87
            TabIndex        =   40
            Top             =   840
            Value           =   -51
            Width           =   2055
         End
         Begin VB.TextBox txt_RX_Sensetivity 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txt_TX_Power 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   360
            Width           =   735
         End
         Begin VB.HScrollBar HS_TxPower 
            Height          =   375
            Left            =   2160
            Max             =   0
            Min             =   -19
            TabIndex        =   37
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "TX Power:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Sensetivity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Label lbl 
         Caption         =   "Tip"
         Height          =   612
         Left            =   240
         TabIndex        =   61
         Top             =   5400
         Width           =   852
      End
      Begin VB.Label Label9 
         Caption         =   "new User section data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5640
         TabIndex        =   59
         Top             =   4800
         Width           =   2532
      End
      Begin VB.Label Label8 
         Caption         =   "old User Section data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2040
         TabIndex        =   58
         Top             =   4800
         Width           =   2292
      End
      Begin VB.Label lbl_EPC_NEW 
         Caption         =   "Tag EPC (New)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   55
         Top             =   2400
         Width           =   1572
      End
      Begin VB.Label Label7 
         Caption         =   "Access PW(New)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   54
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Kill PW ( New)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   53
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Access PW"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   50
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl_TID 
         Caption         =   "TID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   44
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Label lb_Reserved 
         Caption         =   "Reserved"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl_User 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   1572
      End
      Begin VB.Label lb_EPC_data_W 
         Caption         =   "Tag EPC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label lb_TagPW 
         Caption         =   "Kill PW"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   240
         X2              =   9600
         Y1              =   5280
         Y2              =   5280
      End
   End
   Begin VB.Shape TagReadResult 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   8640
      Shape           =   3  '圓形
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lbl_TagAccessResult 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   64
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl_ReadWrite 
      Caption         =   "Access Result:"
      Height          =   375
      Left            =   9120
      TabIndex        =   63
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape_TagDetect 
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   6120
      Shape           =   3  '圓形
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Tag Detect"
      Height          =   495
      Left            =   6600
      TabIndex        =   62
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Connection Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lbl_TagID 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "PID"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11400
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "VID"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11400
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu M_About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**系統名稱：RFIDUSB-HID通訊調試工具
'**描    述：
'**版    本：V1.0.0
'*************************************************************************

Option Explicit
'*********************************************
'
'*********************************************
Dim bAlertable As Long
Dim Capabilities As HIDP_CAPS
Dim DataString As String
Dim DetailData As Long
Dim DetailDataBuffer() As Byte
Dim DeviceAttributes As HIDD_ATTRIBUTES
Dim DevicePathName As String
Dim DeviceInfoSet As Long
Dim ErrorString As String
Dim EventObject As Long
Dim HIDHandle As Long
Dim HIDOverlapped As OVERLAPPED
Dim LastDevice As Boolean
Dim MyDeviceDetected As Boolean
Dim MyDeviceInfoData As SP_DEVINFO_DATA
Dim MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Dim MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Dim Needed As Long
Dim OutputReportData(7) As Byte
Dim PreparsedData As Long
Dim ReadHandle As Long
Dim Result As Long
Dim Security As SECURITY_ATTRIBUTES
Dim Timeout As Boolean
Dim DriveContion As Boolean
'Const MyVendorID = &H1325
'Const MyProductID = &HC02E
Dim MyVendorID As String
Dim MyProductID As String
Dim gTAG As tTAG
'DriveContion = False
Sub RF_Level_Init(Optional tITEM As String = "", Optional tITEM_Value As String = "00")

    Select Case tITEM
        Case "TX_LEVEL"
            'init Tx Level
            WriteReport ("1e 07 00 69  00 02 00 01 01 15 cb cb  cb cb cb cb")
            ReadReport
            WriteReport ("1f 07 00 68  00 02 00 01 15 " + tITEM_Value + " cb cb  cb cb cb cb")  ' 0b = -11
            ReadReport
    
        Case "RX_LEVEL"
            'init Rx Level 9a 0c 00 04  00 07 00 09 01 cd 00 00  00 00 00 cb
            WriteReport ("9a 0c 00 04  00 07 00 09  01  " + tITEM_Value + "  00 00  00 00 00 cb") 'cd = -51
            ReadReport
        Case Else
        
    End Select
    
    Debug.Print tITEM + " = " + tITEM_Value
End Sub
Function FindTheHid() As Boolean

'Makes a series of API calls to locate the desired HID-class device.
'Returns True if the device is detected, False if not detected.

Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

LastDevice = False
MyDeviceDetected = False

'Values for SECURITY_ATTRIBUTES structure:

Security.lpSecurityDescriptor = 0
Security.bInheritHandle = True
Security.nLength = Len(Security)

'******************************************************************************
'一、獲得HID設備的GUID.
'HidD_GetHidGuid
'Get the GUID for all system HIDs.
'Returns: the GUID in HidGuid.
'The routine doesn't return a value in Result
'but the routine is declared as a function for consistency with the other API calls.
'******************************************************************************

Result = HidD_GetHidGuid(HidGuid)
Call DisplayResultOfAPICall("GetHidGuid")

'Display the GUID.

GUIDString = _
    Hex$(HidGuid.Data1) & "-" & _
    Hex$(HidGuid.Data2) & "-" & _
    Hex$(HidGuid.Data3) & "-"

For Count = 0 To 7

    'Ensure that each of the 8 bytes in the GUID displays two characters.
    
    If HidGuid.Data4(Count) >= &H10 Then
        GUIDString = GUIDString & Hex$(HidGuid.Data4(Count)) & " "
    Else
        GUIDString = GUIDString & "0" & Hex$(HidGuid.Data4(Count)) & " "
    End If
Next Count

'lstResults.AddItem "  系統返回的 GUID號： " & GUIDString

'******************************************************************************
'二、找出所有已連接HID設備：
'SetupDiGetClassDevs
'Returns: a handle to a device information set for all installed devices.
'Requires: the HidGuid returned in GetHidGuid.
'******************************************************************************

DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
Call DisplayResultOfAPICall("SetupDiClassDevs（找所有HID設備）")
DataString = GetDataString(DeviceInfoSet, 32)

'******************************************************************************
'三、列舉每一個HID設備:
'SetupDiEnumDeviceInterfaces
'On return, MyDeviceInterfaceData contains the handle to a
'SP_DEVICE_INTERFACE_DATA structure for a detected device.
'Requires:
'the DeviceInfoSet returned in SetupDiGetClassDevs.
'the HidGuid returned in GetHidGuid.
'An index to specify a device.
'******************************************************************************

'Begin with 0 and increment until no more devices are detected.

MemberIndex = 0

Do
    'The cbSize element of the MyDeviceInterfaceData structure must be set to
    'the structure's size in bytes. The size is 28 bytes.
    
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
    
    Call DisplayResultOfAPICall("SetupDiEnumDeviceInterfaces")
    If Result = 0 Then LastDevice = True
    
    'If a device exists, display the information returned.
    
    If Result <> 0 Then
        
        'lstResults.AddItem "  DeviceInfoSet for device " & "找要的設備#" & CStr(MemberIndex) & ": "
         'list Device info on ListBox
  
        
        '******************************************************************************
        '四、取設備的路徑
        'SetupDiGetDeviceInterfaceDetail
        'Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
        'containing information about a device.
        'To retrieve the information, call this function twice.
        'The first time returns the size of the structure in Needed.
        'The second time returns a pointer to the data in DeviceInfoSet.
        'Requires:
        'A DeviceInfoSet returned by SetupDiGetClassDevs and
        'an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
        '*******************************************************************************
        
        MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        Call DisplayResultOfAPICall("SetupDiGetDeviceInterfaceDetail(取設備路徑)")
        Debug.Print "  (OK to say too small)"
        Debug.Print "  Required buffer size for the data: " & Needed
        
        'Store the structure's size.
        
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        'Use a byte array to allocate memory for
        'the MyDeviceInterfaceDetailData structure
        
        ReDim DetailDataBuffer(Needed)
        
        'Store cbSize in the first four bytes of the array.
        
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        'Call SetupDiGetDeviceInterfaceDetail again.
        'This time, pass the address of the first element of DetailDataBuffer
        'and the returned required buffer size in DetailData.
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
        
        Call DisplayResultOfAPICall(" Result of second call:（第二次調用） ")
        Debug.Print "  MyDeviceInterfaceDetailData.cbSize: " & _
            CStr(MyDeviceInterfaceDetailData.cbSize)
        
        'Convert the byte array to a string.
        
        DevicePathName = CStr(DetailDataBuffer())
        
        'Convert to Unicode.
        
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        
        'Strip cbSize (4 bytes) from the beginning.
        
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        Debug.Print "  Device pathname: "
        Debug.Print "    " & DevicePathName
                
        '******************************************************************************
        '五、取得設備的標示代號:
        'CreateFile
        'Returns: a handle that enables reading and writing to the device.
        'Requires:
        'The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
        '******************************************************************************
    
        HIDHandle = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            Security, _
            OPEN_EXISTING, _
            0&, _
            0)
            
        Call DisplayResultOfAPICall("CreateFile（標示代號）")
        Debug.Print "  Returned handle: " & Hex$(HIDHandle) & "h"
        
        'Now we can find out if it's the device we're looking for.
        
        '******************************************************************************
        '取得廠商與產品ID：
        'HidD_GetAttributes
        'Requests information from the device.
        'Requires: The handle returned by CreateFile.
        'Returns: an HIDD_ATTRIBUTES structure containing
        'the Vendor ID, Product ID, and Product Version Number.
        'Use this information to determine if the detected device
        'is the one we're looking for.
        '******************************************************************************
        
        'Set the Size property to the number of bytes in the structure.
        
        DeviceAttributes.Size = LenB(DeviceAttributes)
        Result = HidD_GetAttributes _
            (HIDHandle, _
            DeviceAttributes)
            
        Call DisplayResultOfAPICall("HidD_GetAttributes（取PID,VID）")
        'Call DisplayResultOfAPICall("HidD_GetAttributes（" + MyProductID + "," + MyVendorID + "）")
        If Result <> 0 Then
            Debug.Print "  HIDD_ATTRIBUTES structure filled without error."
        Else
            Debug.Print "  Error in filling HIDD_ATTRIBUTES structure."
        End If
    
        'debug.print  "  Structure size: " & DeviceAttributes.Size
        Debug.Print "  Vendor ID: " & Hex$(DeviceAttributes.VendorID)
        Debug.Print "  Product ID: " & Hex$(DeviceAttributes.ProductID)
        'debug.print  "  Version Number: " & Hex$(DeviceAttributes.VersionNumber)
        
        'Find out if the device matches the one we're looking for.
        
        If (DeviceAttributes.VendorID = MyVendorID) And _
            (DeviceAttributes.ProductID = MyProductID) Then
                
                'It's the desired device.
                
                Debug.Print "  device found!！"
                MyDeviceDetected = True
                DriveContion = True
        Else
                MyDeviceDetected = False
                
                'If it's not the one we want, close its handle.
                
                Result = CloseHandle _
                    (HIDHandle)
                DisplayResultOfAPICall ("CloseHandle（關閉此接口）")
        End If
End If
    
    'Keep looking until we find the device or there are no more left to examine.
    
    MemberIndex = MemberIndex + 1
Loop Until (LastDevice = True) Or (MyDeviceDetected = True)

'Free the memory reserved for the DeviceInfoSet returned by SetupDiGetClassDevs.

Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
Call DisplayResultOfAPICall("DestroyDeviceInfoList（釋放資源）")

If MyDeviceDetected = True Then
    FindTheHid = True
    
    'Learn the capabilities of the device
     
     Call GetDeviceCapabilities
    
    'Get another handle for the overlapped ReadFiles.
    
    ReadHandle = CreateFile _
            (DevicePathName, _
            (GENERIC_READ Or GENERIC_WRITE), _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            Security, _
            OPEN_EXISTING, _
            FILE_FLAG_OVERLAPPED, _
            0)
 
    Call DisplayResultOfAPICall("CreateFile, ReadHandle")
    Debug.Print "  Returned handle: " & Hex$(ReadHandle) & "h"
    Call PrepareForOverlappedTransfer
    'Label3.Caption = "設備聯接成功"
    Label3.Caption = "Connected"
Else
    Debug.Print " device not found。"
    'Label3.Caption = "設備聯接失敗"
    Label3.Caption = "Connect Fail"
End If

End Function

Private Function GetDataString _
    (Address As Long, _
    Bytes As Long) _
As String

'Retrieves a string of length Bytes from memory, beginning at Address.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Offset As Integer
Dim Result$
Dim ThisByte As Byte

For Offset = 0 To Bytes - 1
    Call RtlMoveMemory(ByVal VarPtr(ThisByte), ByVal Address + Offset, 1)
    If (ThisByte And &HF0) = 0 Then
        Result$ = Result$ & "0"
    End If
    Result$ = Result$ & Hex$(ThisByte) & " "
Next Offset

GetDataString = Result$
End Function

Private Function GetErrorString _
    (ByVal LastError As Long) _
As String

'Returns the error message for the last error.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Bytes As Long
Dim ErrorString As String
ErrorString = String$(129, 0)
Bytes = FormatMessage _
    (FORMAT_MESSAGE_FROM_SYSTEM, _
    0&, _
    LastError, _
    0, _
    ErrorString$, _
    128, _
    0)
    
'Subtract two characters from the message to strip the CR and LF.

If Bytes > 2 Then
    GetErrorString = Left$(ErrorString, Bytes - 2)
End If

End Function


Private Sub DisplayResultOfAPICall(FunctionName As String)

'Display the results of an API call.

Dim ErrorString As String

Debug.Print ""
ErrorString = GetErrorString(Err.LastDllError)
Debug.Print FunctionName
Debug.Print "  Result = " & ErrorString

'Scroll to the bottom of the list box.

lstResults.ListIndex = lstResults.ListCount - 1

End Sub

Private Sub btn_Clear_Click()
lstResults.Clear
frmMain.TextIR.Text = ""
End Sub

Private Sub btn_Close_Click()
    Call Shutdown
End Sub

Private Sub BTN_M_Read_Click()
Dim TagScanCount As Integer
'Call ReadReport
'Call ReadReport("ReadTag")
'WriteReport (SelectEPC)
'Sleep (100)
'Call ReadReport
Debug.Print "init RF TX"
'RF_Level_Init

If (0) Then
    WriteReport ("1d 07 00 69  00 02 00 01  01 0a cb cb  cb cb cb cb") '0a
    ReadUSB_Report
    
    WriteReport ("1e 07 00 69  00 02 00 01  01 09 cb cb  cb cb cb cb") '09
    ReadUSB_Report
    
    WriteReport ("1f 07 00 69  00 02 00 01  01 0d cb cb  cb cb cb cb") '0d
    ReadUSB_Report
End If
'Start Scan
WriteReport ("20 11 00 86  00 02 00 00  00 0d 8c 00  05 00 00 01  01 00 01 06  cb cb cb cb  cb cb cb cb  cb cb cb cb")
ReadUSB_Report ("EmptyEPCTextField")

gTAG.TagCOUNT = 0
TagScanCount = 0

Me.Shape_TagDetect.FillColor = &H0& 'black
While ((gTAG.TagCOUNT < 10) And (TagScanCount < 10))
    ReadUSB_Report ("ScanNewTag")
    TagScanCount = TagScanCount + 1
Wend

'StopScan
WriteReport ("21 0a 00 8c  00 05 00 00  01 00 00 00  00 cb cb cb")
ReadUSB_Report


End Sub
Function SelectEPC()
    Dim vSelectEPC As String
    vSelectEPC = READER_T_EPC + frmMain.txt_EPC_Data.Text + "AA"
    Debug.Print vSelectEPC
    SelectEPC = vSelectEPC
End Function

Private Sub BTN_M_ReadEPC_Click()

    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_EPC)
        ReadUSB_Report ("ReadEPC")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_ReadReserved_Click()
    
    
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_RSV)
        ReadUSB_Report ("ReadReserved")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
    
End Sub

Private Sub BTN_M_ReadTID_Click()
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_TID)
        ReadUSB_Report ("ReadTID")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
    

End Sub

Private Sub BTN_M_ReadUser_Click()
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_USR)
        ReadUSB_Report ("ReadUser")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_Write_Click()
    WriteReport (SelectEPC)
    ReadUSB_Report
    Debug.Print "Select Tag'EPC for Writing prepare"
End Sub

Private Sub BTN_M_WriteEPC_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    gTAG.DATAFORMAT_ERR = False
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00") And (gTAG.DATAFORMAT_ERR = False))
        WriteReport (SelectEPC)
        ReadUSB_Report
        DataStr = READER_W_EPC + "00000000" + Me.txt_EPC_Data_New.Text
        Debug.Print DataStr
        WriteReport (DataStr)
        Sleep (30)
        ReadUSB_Report ("ReadEPC_AfterWrote")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_WriteReservedAccess_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        'DataStr = READER_W_RSV_ACCESS_PW + Me.txt_Reserved_AccessPW.Text + Me.txt_Reserved_AccessPW_New.Text
        DataStr = READER_W_RSV_ACCESS_PW + "00000000" + Me.txt_Reserved_AccessPW_New.Text
        WriteReport (DataStr)
        ReadUSB_Report ("ReadReservedAfterUpdateAccessPW")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_WriteReservedKill_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        'DataStr = READER_W_RSV_ACCESS_PW + Me.txt_Reserved_AccessPW.Text + Me.txt_Reserved_AccessPW_New.Text
        DataStr = READER_W_RSV_KILL_PW + "00000000" + Me.txt_Reserved_KillPW_New.Text
        WriteReport (DataStr)
        ReadUSB_Report ("ReadReservedAfterUpdateKillPW")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_WriteUser_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    'MsgBox "UserNewLength =" & Len(Me.txt_User_New.Text)
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        DataStr = READER_W47B_USR + "00000000" + Mid(Me.txt_User_New.Text, 1, 94)
        WriteReport (DataStr)
        Sleep (30)
        
        DataStr = READER_W17B_USR + Mid(Me.txt_User_New.Text, 95, 34)
        WriteReport (DataStr)
        
        ReadUSB_Report ("ReadUser_AfterWrote64B")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub btn_RF_init_Click()
    'init Tx Level
  If (1) Then 'Tx : -11 , Rx : -51
    Call RF_Level_Init("TX_LEVEL", "0B")
    Call RF_Level_Init("RX_LEVEL", "CD")
    Me.HS_TxPower.Value = -11
    Me.HS_RxSensetivity.Value = -51
  Else         'Tx : 0  , Rx : -70
    Call RF_Level_Init("TX_LEVEL", "01")
    Call RF_Level_Init("RX_LEVEL", "BA")
    Me.HS_TxPower.Value = -1
    Me.HS_RxSensetivity.Value = -70
  
  End If
End Sub

Private Sub btn_ScanStart_Click()
WriteReport (READER_T_STARTSCAN)
ReadReport
frmMain.Timer1.Enabled = True
End Sub

Private Sub btn_ScanStop_Click()
WriteReport (READER_T_STOPSCAN)
ReadReport
frmMain.Timer1.Enabled = False
End Sub

Private Sub btn_Test_Click()
WriteReport (READER_T_TEST)
ReadReport
End Sub

Private Sub Clear_Click()

End Sub

Private Sub Command1_Click()
'test command

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub btn_Timer1_Dis_Click()
frmMain.Timer1.Enabled = False
End Sub

Private Sub btn_Timer1_En_Click()
frmMain.Timer1.Enabled = True
End Sub



Private Sub contion_Click()

MyVendorID = "&H" & (Trim(TextVID.Text))   'Hex(temp_i)
MyProductID = "&H" & (Trim(TextPID.Text))

MyVendorID = "&H1325"
MyProductID = -16338

If MyVendorID = "0" And MyProductID = "0" Then
MsgBox ("please input VendorID and  ProductID")
Exit Sub
End If
FindTheHid




    
End Sub



Private Sub Form_Load()
frmMain.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Shutdown
End Sub

Private Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'Returns: a pointer to a buffer containing information about the device's capabilities.
'Requires: A handle returned by CreateFile.
'There's no need to access the buffer directly,
'but HidP_GetCaps and other API functions require a pointer to the buffer.
'******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant

'Preparsed Data is a pointer to a routine-allocated buffer.

Result = HidD_GetPreparsedData _
    (HIDHandle, _
    PreparsedData)
Call DisplayResultOfAPICall("HidD_GetPreparsedData")

'Copy the data at PreparsedData into a byte array.

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)
Call DisplayResultOfAPICall("RtlMoveMemory")

ppDataString = ppData()

'Convert the data to Unicode.

ppDataString = StrConv(ppDataString, vbUnicode)

'******************************************************************************
'HidP_GetCaps
'Find out the device's capabilities.
'For standard devices such as joysticks, you can find out the specific
'capabilities of the device.
'For a custom device, the software will probably know what the device is capable of,
'so this call only verifies the information.
'Requires: The pointer to a buffer containing the information.
'The pointer is returned by HidD_GetPreparsedData.
'Returns: a Capabilites structure containing the information.
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)

Call DisplayResultOfAPICall("HidP_GetCaps")
Debug.Print "  Last error: " & ErrorString
Debug.Print "  Usage: " & Hex$(Capabilities.Usage)
Debug.Print "  Usage Page: " & Hex$(Capabilities.UsagePage)
Debug.Print "  Input Report Byte Length: " & Capabilities.InputReportByteLength
Debug.Print "  Output Report Byte Length: " & Capabilities.OutputReportByteLength
Debug.Print "  Feature Report Byte Length: " & Capabilities.FeatureReportByteLength
Debug.Print "  Number of Link Collection Nodes: " & Capabilities.NumberLinkCollectionNodes
Debug.Print "  Number of Input Button Caps: " & Capabilities.NumberInputButtonCaps
Debug.Print "  Number of Input Value Caps: " & Capabilities.NumberInputValueCaps
Debug.Print "  Number of Input Data Indices: " & Capabilities.NumberInputDataIndices
Debug.Print "  Number of Output Button Caps: " & Capabilities.NumberOutputButtonCaps
Debug.Print "  Number of Output Value Caps: " & Capabilities.NumberOutputValueCaps
Debug.Print "  Number of Output Data Indices: " & Capabilities.NumberOutputDataIndices
Debug.Print "  Number of Feature Button Caps: " & Capabilities.NumberFeatureButtonCaps
Debug.Print "  Number of Feature Value Caps: " & Capabilities.NumberFeatureValueCaps
Debug.Print "  Number of Feature Data Indices: " & Capabilities.NumberFeatureDataIndices

'******************************************************************************
'HidP_GetValueCaps
'Returns a buffer containing an array of HidP_ValueCaps structures.
'Each structure defines the capabilities of one value.
'This application doesn't use this data.
'******************************************************************************

'This is a guess. The byte array holds the structures.

Dim ValueCaps(1023) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
   
Call DisplayResultOfAPICall("HidP_GetValueCaps")

'debug.print  "ValueCaps= " & GetDataString((VarPtr(ValueCaps(0))), 180)
'To use this data, copy the byte array into an array of structures.

'Free the buffer reserved by HidD_GetPreparsedData

Result = HidD_FreePreparsedData _
    (PreparsedData)
Call DisplayResultOfAPICall("HidD_FreePreparsedData")

End Sub


Private Sub PrepareForOverlappedTransfer()

'******************************************************************************
'CreateEvent
'Creates an event object for the overlapped structure used with ReadFile.
'Requires a security attributes structure or null,
'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),
'Initial state = True (signaled),
'and event object name (optional)
'Returns a handle to the event object.
'******************************************************************************

If EventObject = 0 Then
    EventObject = CreateEvent _
        (Security, _
        True, _
        True, _
        "")
End If
    
Call DisplayResultOfAPICall("CreateEvent")
    
'Set the members of the overlapped structure.

HIDOverlapped.Offset = 0
HIDOverlapped.OffsetHigh = 0
HIDOverlapped.hEvent = EventObject
End Sub



Private Sub Shutdown()

'Actions that must execute when the program ends.

'Close the open handles to the device.

Result = CloseHandle _
    (HIDHandle)
Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")

Result = CloseHandle _
    (ReadHandle)
Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub HS_RxSensetivity_Change()
    Dim RxLevel As String
    RxLevel = Hex$(CStr(256 + frmMain.HS_RxSensetivity.Value))
    
    frmMain.txt_RX_Sensetivity.Text = frmMain.HS_RxSensetivity.Value
    Call RF_Level_Init("RX_LEVEL", RxLevel)
End Sub

Private Sub HS_TxPower_Change()
    Dim TxPower As String
    Dim bb As Byte
    'TxPower = CStr(0 - frmMain.HS_TxPower.Value)
    TxPower = Hex$(CStr(0 - frmMain.HS_TxPower.Value))
    frmMain.txt_TX_Power.Text = frmMain.HS_TxPower.Value
    If (Len(TxPower) < 2) Then
        TxPower = "0" + TxPower
    End If
    Call RF_Level_Init("TX_LEVEL", TxPower) ' Min(19) ~ Max(0)


End Sub

Private Sub Iread_Click()
Call ReadReport
End Sub
Private Sub ReadReport(Optional RdCmdStr As String = "")

    'Read data from the device.
    
    Dim Count
    Dim NumberOfBytesRead As Long
    
    'Allocate a buffer for the report.
    'Byte 0 is the report ID.
    
    Dim ReadBuffer() As Byte
    Dim UBoundReadBuffer As Integer
    
    '******************************************************************************
    'ReadFile
    'Returns: the report in ReadBuffer.
    'Requires: a device handle returned by CreateFile
    '(for overlapped I/O, CreateFile must be called with FILE_FLAG_OVERLAPPED),
    'the Input report length in bytes returned by HidP_GetCaps,
    'and an overlapped structure whose hEvent member is set to an event object.
    '******************************************************************************
    
    Dim ByteValue As String
    Dim ByteValueOutput As String
    
    If MyDeviceDetected = False And DriveContion = False Then
          Debug.Print "讀出失敗請先聯接"
          Exit Sub
          End If
          
    
          
    'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
    
    ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)
    
    'Scroll to the bottom of the list box.
    
    lstResults.ListIndex = lstResults.ListCount - 1
    
    'Do an overlapped ReadFile.
    'The function returns immediately, even if the data hasn't been received yet.
    
    Result = ReadFile _
        (ReadHandle, _
        ReadBuffer(0), _
        CLng(Capabilities.InputReportByteLength), _
        NumberOfBytesRead, _
        HIDOverlapped)
    Call DisplayResultOfAPICall("ReadFile")
    
    Debug.Print "waiting for ReadFile"
    
    'Scroll to the bottom of the list box.
    
    lstResults.ListIndex = lstResults.ListCount - 1
    bAlertable = True
    
    '******************************************************************************
    'WaitForSingleObject
    'Used with overlapped ReadFile.
    'Returns when ReadFile has received the requested amount of data or on timeout.
    'Requires an event object created with CreateEvent
    'and a timeout value in milliseconds.
    '******************************************************************************
    Result = WaitForSingleObject _
        (EventObject, _
        100)
    Call DisplayResultOfAPICall("WaitForSingleObject")
    
    'Find out if ReadFile completed or timeout.
    
    Select Case Result
        Case WAIT_OBJECT_0
            
            'ReadFile has completed
            
            Debug.Print "ReadFile completed successfully."
        Case WAIT_TIMEOUT
            
            'Timeout
            
            Debug.Print "Readfile timeout"
            
            'Cancel the operation
            
            '*************************************************************
            'CancelIo
            'Cancels the ReadFile
            'Requires the device handle.
            'Returns non-zero on success.
            '*************************************************************
            Result = CancelIo _
                (ReadHandle)
            Debug.Print "************ReadFile timeout*************"
            Debug.Print "CancelIO"
            Call DisplayResultOfAPICall("CancelIo")
            
            'The timeout may have been because the device was removed,
            'so close any open handles and
            'set MyDeviceDetected=False to cause the application to
            'look for the device on the next attempt.
            
            'CloseHandle (HIDHandle)
            'Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")
            'CloseHandle (ReadHandle)
            'Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")
            'MyDeviceDetected = False
        Case Else
            Debug.Print "Readfile undefined error"
            MyDeviceDetected = False
    End Select
        
    Debug.Print " Report ID: " & ReadBuffer(0)
    Debug.Print " Report Data:"
    
    'frmMain.TextIR.Text = ""
    'ByteValueOutput = ""
    
    For Count = 1 To UBound(ReadBuffer)
        
        'Add a leading 0 to values 0 - Fh.
        
        If Len(Hex$(ReadBuffer(Count))) < 2 Then
            ByteValue = "0" & Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        Else
            ByteValue = Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        End If
        
        'If (Len(ByteValueOutput) Mod 128) = 0 Then
            'ByteValueOutput = ByteValueOutput + vbCrLf
            'ByteValueOutput = ByteValueOutput + Chr(13) + Chr(10)
        'End If
    
    Next Count
    

    'print log to TextIR text box
    
    
        
        Select Case RdCmdStr
          Case "ReadTag"
                    gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                    frmMain.txt_EPC_Data.Text = gTAG.EPC

          Case Else
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
        End Select
    
    '******************************************************************************
    'ResetEvent
    'Sets the event object in the overlapped structure to non-signaled.
    'Requires a handle to the event object.
    'Returns non-zero on success.
    '******************************************************************************
    
    Call ResetEvent(EventObject)
    Call DisplayResultOfAPICall("ResetEvent")

End Sub
Private Sub ReadUSB_Report(Optional RdCmdStr As String = "")

    'Read data from the device.
    
    Dim Count
    Dim NumberOfBytesRead As Long
    
    'Allocate a buffer for the report.
    'Byte 0 is the report ID.
    
    Dim ReadBuffer() As Byte
    Dim UBoundReadBuffer As Integer
    
    '******************************************************************************
    'ReadFile
    'Returns: the report in ReadBuffer.
    'Requires: a device handle returned by CreateFile
    '(for overlapped I/O, CreateFile must be called with FILE_FLAG_OVERLAPPED),
    'the Input report length in bytes returned by HidP_GetCaps,
    'and an overlapped structure whose hEvent member is set to an event object.
    '******************************************************************************
    
    Dim ByteValue As String
    Dim ByteValueOutput As String
    
    Dim TagReadResult As String '標籤讀取結果
    Debug.Print RdCmdStr
    
    If MyDeviceDetected = False And DriveContion = False Then
          Debug.Print "讀出失敗請先聯接"
          Exit Sub
    End If
          
    
          
    'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
    
    ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)
    
    
    'Do an overlapped ReadFile.
    'The function returns immediately, even if the data hasn't been received yet.
    
    Result = ReadFile _
        (ReadHandle, _
        ReadBuffer(0), _
        CLng(Capabilities.InputReportByteLength), _
        NumberOfBytesRead, _
        HIDOverlapped)
    
    Debug.Print "waiting for ReadFile"
    
    'Scroll to the bottom of the list box.
    
    bAlertable = True
    
    '******************************************************************************
    'WaitForSingleObject
    'Used with overlapped ReadFile.
    'Returns when ReadFile has received the requested amount of data or on timeout.
    'Requires an event object created with CreateEvent
    'and a timeout value in milliseconds.
    '******************************************************************************
    Result = WaitForSingleObject _
        (EventObject, _
        100)
    Call DisplayResultOfAPICall("WaitForSingleObject")
    
    'Find out if ReadFile completed or timeout.
    
    Select Case Result
        Case WAIT_OBJECT_0
            
            'ReadFile has completed
            
            Debug.Print "ReadFile completed successfully."
        Case WAIT_TIMEOUT
            
            'Timeout
            
            Debug.Print "Readfile timeout"
            
            'Cancel the operation
            
            '*************************************************************
            'CancelIo
            'Cancels the ReadFile
            'Requires the device handle.
            'Returns non-zero on success.
            '*************************************************************
            Result = CancelIo _
                (ReadHandle)
            Debug.Print "************ReadFile timeout*************"
            Debug.Print "CancelIO"
            Call DisplayResultOfAPICall("CancelIo")
            
            'The timeout may have been because the device was removed,
            'so close any open handles and
            'set MyDeviceDetected=False to cause the application to
            'look for the device on the next attempt.
            
            'CloseHandle (HIDHandle)
            'Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")
            'CloseHandle (ReadHandle)
            'Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")
            'MyDeviceDetected = False
        Case Else
            Debug.Print "Readfile undefined error"
            MyDeviceDetected = False
    End Select
        
    Debug.Print " Report  : Result : " & Result & " ,ID: " & ReadBuffer(0)
    'Debug.Print " Report Data:"
    
    'frmMain.TextIR.Text = ""
    'ByteValueOutput = ""
    
    For Count = 1 To UBound(ReadBuffer)
        
        'Add a leading 0 to values 0 - Fh.
        
        If Len(Hex$(ReadBuffer(Count))) < 2 Then
            ByteValue = "0" & Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        Else
            ByteValue = Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        End If
    Next Count
    
        TagReadResult = "00"
        Me.TagReadResult.FillColor = &HC0C0C0
    
        Select Case RdCmdStr
          Case "EmptyEPCTextField"
                    frmMain.txt_EPC_Data.Text = ""
          Case "ReadReserved"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    frmMain.txt_Reserved_AccessPW.Text = ""
                    frmMain.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        'frmMain.txt_Reserved_KillPW.Text = Mid(ByteValueOutput, 17, 8)
                        'frmMain.txt_Reserved_AccessPW.Text = Mid(ByteValueOutput, 25, 8)
                        frmMain.txt_Reserved_KillPW.Text = Mid(ByteValueOutput, 17, 8)
                        frmMain.txt_Reserved_AccessPW.Text = Mid(ByteValueOutput, 25, 8)
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadReservedAfterUpdateAccessPW"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'frmMain.txt_Reserved_AccessPW.Text = ""
                    'frmMain.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        frmMain.txt_Reserved_AccessPW.Text = frmMain.txt_Reserved_AccessPW_New
                        'frmMain.txt_Reserved_KillPW.Text = frmMain.txt_Reserved_KillPW_New.Text
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadReservedAfterUpdateKillPW"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'frmMain.txt_Reserved_AccessPW.Text = ""
                    'frmMain.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        'frmMain.txt_Reserved_AccessPW.Text = frmMain.txt_Reserved_AccessPW_New
                        frmMain.txt_Reserved_KillPW.Text = frmMain.txt_Reserved_KillPW_New.Text
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
          Case "ReadEPC"
                    '2816000800DE001193B23400E2003000390701110610D48103055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'frmMain.txt_EPC_Data = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.EPC = Mid(ByteValueOutput, 25, 24)
                        frmMain.txt_EPC_Data = gTAG.EPC
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        
                    End If
                    Debug.Print "Tag EPC Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadEPC_AfterWrote"
                    '2816000800DE001193B23400E2003000390701110610D48103055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'frmMain.txt_EPC_Data = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        'f30700070000000206a5bbbb11111111033400e2003000390701110610d4827c

                        gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                        frmMain.txt_EPC_Data = frmMain.txt_EPC_Data_New
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "ByteValueOutput = " + ByteValueOutput
                        'Debug.Print "Result check:" & Mid(ByteValueOutput, 3, 8)
                        'Can't just confirm  byte_6, also confirm byte2 , byte4 should as 07 07
                        If (Mid(ByteValueOutput, 3, 8) <> "07000700") Then
                            gTAG.ACC_RESULT = "FF"
                        End If
                    Else
                        
                    End If
                    Debug.Print "Tag EPC Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          Case "ReadTID"
                    '2D1E000800DE0019E2003412012EF8000686D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    frmMain.txt_TID = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.TID = Mid(ByteValueOutput, 17, 48)
                        frmMain.txt_TID = gTAG.TID
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        gTAG.TID = Mid(ByteValueOutput, 17, 48)
                        frmMain.txt_TID = gTAG.TID
                        Me.TagReadResult.FillColor = &HFF00&
                    End If
                    Debug.Print "Tag TID Result=" + gTAG.ACC_RESULT
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          Case "ReadUser"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    frmMain.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.USER = Mid(ByteValueOutput, 17, 112)
                        frmMain.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "Mid(ByteValueOutput, 15, 2)= " + Mid(ByteValueOutput, 15, 2)
                        gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 15, 2))
                        ReadUSB_Report ("ReadUserPart2")
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadUserPart2"
                    gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 3, 2))
                    If (gTAG.DATA_LENGTH > 0) Then
                        'gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 3, 2))
                        gTAG.USER = gTAG.USER + Mid(ByteValueOutput, 7, 16)
                        frmMain.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
          Case "ReadUser_AfterWrote"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    frmMain.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.USER = Mid(ByteValueOutput, 17, 112)
                        frmMain.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "Mid(ByteValueOutput, 15, 2)= " + Mid(ByteValueOutput, 15, 2)
                        gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 15, 2))
                        ReadUSB_Report ("ReadUserPart2")
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadUser_AfterWrote64B"
                    'frmMain.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    
                    If (Mid(ByteValueOutput, 3, 8) = "07000700") Then
                        gTAG.USER = frmMain.txt_User_New
                        frmMain.txt_User = frmMain.txt_User_New
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        gTAG.ACC_RESULT = "FF"
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
                    
          Case "ScanNewTag"
                    '0F1C0005000000170100018FDA16D20D0E3400E2003000390701110610D481000000000000000000000000000000000000000000000000000000000000000000
                    'gTAG.TagCOUNT = CInt(Mid(ByteValueOutput, 21, 2))
                    
                    
                    If (Mid(ByteValueOutput, 7, 2) = "05") Then
                        If (CInt(Mid(ByteValueOutput, 21, 2)) > 0) Then
                             gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                             frmMain.txt_EPC_Data.Text = gTAG.EPC
                             frmMain.txt_EPC_Data_New.Text = gTAG.EPC
                             Me.Shape_TagDetect.FillColor = &HFF00& ' green
                         End If
                    Else
                            
                    End If
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = TagReadResult

          Case Else
                    frmMain.TextIR.Text = frmMain.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = ""
        End Select
    
    '******************************************************************************
    'ResetEvent
    'Sets the event object in the overlapped structure to non-signaled.
    'Requires a handle to the event object.
    'Returns non-zero on success.
    '******************************************************************************
    
    Call ResetEvent(EventObject)
    Call DisplayResultOfAPICall("ResetEvent")

End Sub


Private Sub Iwrite_Click()

End Sub

Private Sub WriteReport(Optional TxStr As String = "")

'Send data to the device.

Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim ReadBuffer() As Byte
Dim SendBuffer() As Byte
Dim temp As String
Dim Cwritlong As Integer
Dim Wcont As Integer
Dim SendTxStr As String
'******************************************************************************
'WriteFile
'Sends a report to the device.
'Returns: success or failure.
'Requires: the handle returned by CreateFile and
'The output report byte length returned by HidP_GetCaps
'******************************************************************************

If MyDeviceDetected = False And DriveContion = False Then
      Debug.Print "device write fail , please check device is plugin ready!"
      Exit Sub
    ElseIf MyDeviceDetected = False And DriveContion = True Then
    MyDeviceDetected = FindTheHid
    End If
    
If MyDeviceDetected = True Then
'The SendBuffer array begins at 0, so subtract 1 from the number of bytes.
ReDim SendBuffer(Capabilities.OutputReportByteLength - 1)

'temp = TextIW.Text
temp = TxStr
temp = Replace(temp, " ", "")
Cwritlong = Len(temp) / 2
If Cwritlong < Capabilities.OutputReportByteLength - 1 Then
    For Wcont = 1 To Capabilities.OutputReportByteLength - 1 - Cwritlong
    temp = temp + "00"
    Next Wcont
End If
frmMain.TextIR.Text = frmMain.TextIR.Text + "> " + temp + vbCrLf
'The first byte is the Report ID

SendBuffer(0) = 0

'The next bytes are data
On Error GoTo ERROR_Handle

For Count = 0 To Capabilities.OutputReportByteLength - 2
    '從文本框中取出數放到發送中
    SendBuffer(Count + 1) = "&H" & Trim(Mid(temp, Count * 2 + 1, 2))
Next Count

NumberOfBytesWritten = 0

Result = WriteFile _
    (HIDHandle, _
    SendBuffer(0), _
    CLng(Capabilities.OutputReportByteLength), _
    NumberOfBytesWritten, _
    0)
Call DisplayResultOfAPICall("WriteFile")

Debug.Print " OutputReportByteLength = " & Capabilities.OutputReportByteLength
Debug.Print " NumberOfBytesWritten = " & NumberOfBytesWritten
Debug.Print " Report ID: " & SendBuffer(0)
Debug.Print " Report Data:"

For Count = 1 To UBound(SendBuffer)
    Debug.Print Count & " " & Hex$(SendBuffer(Count))
    'SendTxStr = SendTxStr + Hex$(SendBuffer(Count))
Next Count
    'Debug.Print SendTxStr
End If

Exit Sub

ERROR_Handle:
    'MsgBox Err.Number & "Data Error , Please make sure data match hex format 00 ~ FF  !"
    If Err.Number = 13 Then
        lstResults.AddItem "Data Error , Please make sure input data match Hex format 00 ~ FF  !"
        gTAG.DATAFORMAT_ERR = True ' set gTAG.DATAFORMAT_ERR = True for exit write loop
    Else
        lstResults.AddItem "Err Code : " & Err.Number & "  !"
    End If
End Sub

Private Sub TextIW_Change()

End Sub



Private Sub M_About_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
    frmMain.TextIR.Text = ""
    ReadReport
End Sub

Function Base64Encode(Str() As Byte) As String
On Error GoTo over
Dim buf() As Byte, length As Long, mods As Long
Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
mods = (UBound(Str) + 1) Mod 3
length = UBound(Str) + 1 - mods
ReDim buf(length / 3 * 4 + IIf(mods <> 0, 3, 0))
Dim i As Long
For i = 0 To length - 1 Step 3
buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
Next
If mods = 1 Then
buf(length / 3 * 4) = (Str(length) And &HFC) / &H4
buf(length / 3 * 4 + 1) = (Str(length) And &H3) * &H10
buf(length / 3 * 4 + 2) = 64
buf(length / 3 * 4 + 3) = 64
ElseIf mods = 2 Then
buf(length / 3 * 4) = (Str(length) And &HFC) / &H4
buf(length / 3 * 4 + 1) = (Str(length) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
buf(length / 3 * 4 + 2) = (Str(length) And &HF) * &H4
buf(length / 3 * 4 + 3) = 64
End If
For i = 0 To UBound(buf)
Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
Next
over:
End Function

Function Base64Uncode(B64 As String) As Byte()
On Error GoTo over
Dim OutStr() As Byte, i As Long, j As Long
Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)
Dim length As Long, mods As Long
mods = Len(B64) Mod 4
length = Len(B64) - mods
ReDim OutStr(length / 4 * 3 - 1 + Switch(mods = 2, 1, mods = 3, 2))
For i = 1 To length Step 4
Dim buf(3) As Byte
For j = 0 To 3
buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1
Next
OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
Next
If mods = 2 Then
OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
ElseIf mods = 3 Then
OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
OutStr(length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 3, 1)) - 1) And &H3C) / &H4
End If
Base64Uncode = OutStr
over:
End Function

