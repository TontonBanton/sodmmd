VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMIS 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MATERIALS ISSUANCE "
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   Icon            =   "MIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameItemAsset 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1140
      Left            =   14160
      TabIndex        =   60
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtAssetAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   5950
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   960
         Width           =   2200
      End
      Begin VB.TextBox txtAssetPO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   240
         TabIndex        =   70
         Top             =   3525
         Width           =   3700
      End
      Begin VB.CommandButton cmdFixAsset 
         Caption         =   "TO FIXED ASSETS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4350
         Width           =   3705
      End
      Begin VB.TextBox txtAssetAvail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtAssetATF 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   4440
         TabIndex        =   67
         Top             =   2200
         Width           =   3700
      End
      Begin VB.TextBox txtAssetSerial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   240
         TabIndex        =   66
         Top             =   2880
         Width           =   3700
      End
      Begin VB.TextBox txtAssetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   3800
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   960
         Width           =   2100
      End
      Begin VB.TextBox txtAssetUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtAssetName 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   250
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   360
         Width           =   7900
      End
      Begin VB.TextBox txtAssetTag 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   4440
         TabIndex        =   62
         Top             =   3525
         Width           =   3700
      End
      Begin VB.ComboBox txtAssetGroup 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "MIS.frx":1E72
         Left            =   240
         List            =   "MIS.frx":1E85
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2200
         Width           =   3735
      End
      Begin MSMask.MaskEdBox txtATFDate 
         Height          =   600
         Left            =   4440
         TabIndex        =   72
         Top             =   2880
         Width           =   3700
         _ExtentX        =   6535
         _ExtentY        =   1058
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "COST"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   13
         Left            =   5320
         TabIndex        =   74
         Top             =   1605
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   7230
         TabIndex        =   73
         Top             =   1605
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   850
      Left            =   0
      TabIndex        =   36
      Top             =   9800
      Width           =   11775
      Begin VB.CommandButton cmdNew 
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   9400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   2430
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4750
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14760
      TabIndex        =   58
      Top             =   0
      Width           =   2175
   End
   Begin VB.Frame frameItemAdd 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   7600
      Left            =   720
      TabIndex        =   41
      Top             =   1880
      Visible         =   0   'False
      Width           =   7045
      Begin VB.ComboBox txtMISSeason 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1EF3
         Left            =   5000
         List            =   "MIS.frx":1EFD
         TabIndex        =   23
         Text            =   "-"
         Top             =   5040
         Width           =   1620
      End
      Begin VB.ComboBox txtMISAsset 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F11
         Left            =   5000
         List            =   "MIS.frx":1F13
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4560
         Width           =   1620
      End
      Begin VB.ComboBox txtMISBlock 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F15
         Left            =   5000
         List            =   "MIS.frx":1F17
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   5520
         Width           =   1620
      End
      Begin VB.TextBox txtMISAvail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   700
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1950
         Width           =   2350
      End
      Begin VB.TextBox txtMISAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   700
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2730
         Width           =   4700
      End
      Begin VB.TextBox txtMISUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   700
         Left            =   4280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1200
         Width           =   2350
      End
      Begin VB.ComboBox txtMISEquip 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F19
         Left            =   1920
         List            =   "MIS.frx":1F1B
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   6000
         Width           =   4695
      End
      Begin VB.ComboBox txtMISArea 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F1D
         Left            =   1920
         List            =   "MIS.frx":1F1F
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   5520
         Width           =   3100
      End
      Begin VB.ComboBox txtMISTransact 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F21
         Left            =   1920
         List            =   "MIS.frx":1F23
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5040
         Width           =   3135
      End
      Begin VB.ComboBox txtMISDept 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "MIS.frx":1F25
         Left            =   1920
         List            =   "MIS.frx":1F38
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4560
         Width           =   3135
      End
      Begin VB.CommandButton cmdSaveMISDetails 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6720
         Width           =   4695
      End
      Begin VB.TextBox txtMISCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   700
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1200
         Width           =   2350
      End
      Begin VB.TextBox txtMISItem 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   600
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Width           =   6270
      End
      Begin VB.TextBox txtMISGroup 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   4065
         Width           =   4700
      End
      Begin VB.TextBox txtMISQty 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4280
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1950
         Width           =   2350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "U/COST"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   480
         TabIndex        =   59
         Top             =   1400
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   480
         TabIndex        =   57
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "EQUIPMENT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   300
         TabIndex        =   55
         Top             =   6120
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   300
         TabIndex        =   54
         Top             =   5640
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   300
         TabIndex        =   53
         Top             =   5160
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "GROUP"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   300
         TabIndex        =   52
         Top             =   4155
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "CHARGING:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   300
         TabIndex        =   51
         Top             =   4680
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   300
         TabIndex        =   49
         Top             =   5640
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "INV / QTY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   48
         Top             =   2160
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "EDIT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame frameItems 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1635
      Left            =   8400
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   3120
      Begin VB.CommandButton cmdAddClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   11350
         TabIndex        =   50
         Top             =   240
         Width           =   3700
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "F1 - ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   8040
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   3500
      End
      Begin VB.TextBox txtItemSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   150
         TabIndex        =   11
         Top             =   260
         Width           =   11000
      End
      Begin MSComctlLib.ListView lvwMISItems 
         Height          =   5445
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   9604
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame frameSearch 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1410
      Left            =   8400
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   3225
      Begin VB.ComboBox cboMISSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "MIS.frx":1F68
         Left            =   150
         List            =   "MIS.frx":1F78
         TabIndex        =   28
         Top             =   250
         Width           =   2805
      End
      Begin VB.TextBox txtMISSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3050
         MaxLength       =   50
         TabIndex        =   29
         Top             =   240
         Width           =   3525
      End
      Begin MSComctlLib.ListView lvwMISSearch 
         Height          =   1365
         Left            =   0
         TabIndex        =   44
         Top             =   720
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   2408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483637
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtMISTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   39
      Top             =   9800
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1140
      Left            =   0
      TabIndex        =   31
      Top             =   800
      Width           =   17055
      Begin VB.TextBox txtMISRemarks 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7560
         TabIndex        =   10
         Top             =   600
         Width           =   9075
      End
      Begin VB.TextBox txtMISNum 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   6
         Top             =   200
         Width           =   2500
      End
      Begin VB.TextBox txtMISRequest 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7560
         TabIndex        =   9
         Top             =   200
         Width           =   9075
      End
      Begin MSMask.MaskEdBox txtMISDate 
         Height          =   420
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMISLocation 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   7
         Top             =   200
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   6000
         TabIndex        =   56
         Top             =   670
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "MIS NUMBER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   250
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "MIS DATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   255
         TabIndex        =   33
         Top             =   670
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "RECEIVED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   6000
         TabIndex        =   32
         Top             =   250
         Width           =   1230
      End
   End
   Begin MSComctlLib.ListView lvwMIS 
      Height          =   7755
      Left            =   0
      TabIndex        =   35
      Top             =   1920
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   13679
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   120
      Picture         =   "MIS.frx":1F99
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2085
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   6843
      TabIndex        =   43
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   6843
      TabIndex        =   42
      Top             =   435
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   1
      Left            =   12000
      TabIndex        =   40
      Top             =   10080
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   5938
      Picture         =   "MIS.frx":2B89
      Stretch         =   -1  'True
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "FormMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset

Private MISId, EncodeMode, ItemName, ItemGroup, MISResult, MISDetails               As String
Private strsql, ItemCode, ItemNoDel, TxtVal, NumVal, AvailStock                     As String
Private MISRef, MSSItemNo, MISTotal, MISAvail, TransactID, ItemInvID                As Double
Private ItemAmount, ItemBal, ItemCost, InvAmount, ItemQtyDel                        As Double
Dim AssetID, AssetGroup, AssetATF   As String
Private LI            As ListItem
Private Sub Command1_Click()
'strsql = "Update MISDetails SET MISItemNo = 1"
'CommandExecute
'strsql = "Update MISDetails2 SET MISBlock = MISRemark"
'CommandExecute
'strsql = "Update MISDetails2 SET MISInfoDetails =  MISTransact & '" & "-" & "' & MISPhase & '" & "-" & "' & MISEquip "
'CommandExecute
'strsql = "Update Materials SET Unit = '" & "ROL" & "' WHERE Unit like '" & "ROLL" & "'"
'CommandExecute
'strsql = "Update MISDetails SET MISPhase = '" & "POM PH-1" & "' WHERE MISPhase like '" & "POM PHASE-1" & "'"
'CommandExecute
End Sub

'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwMIS
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
 Private Sub GetFrom()
   mmsAdoCmd.CommandText = "Select * from Settings"
   Set mmsADORst = mmsAdoCmd.Execute
   txtMISLocation = mmsADORst.Fields("Area") & "-" & mmsADORst.Fields("AreaLoc")
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     'FormMainMenu.cmdExit.SetFocus
End Sub
Private Sub CommandExecute()
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdNew_Click()
Dim MISGroup As String
    EncodeMode = "A": BoxState True: ClearBox: ClearItemBox: ClearFrame
    ButtonState False: cmdAdd.Enabled = True: cmdCancel.Enabled = True

      lvwMIS.ListItems.Clear
      MISGroup = FormMainMenu.lblArea.Caption & "-" & "MIS"
      MISId = Format$(GetNextMISID, "000000")
      txtMISNum.Text = MISGroup & "-" & MISId
      MSSItemNo = 0
      txtMISDate = Format$(Now, "mm/dd/yyyy")
      txtMISDate.SetFocus
      GetFrom
      DeleteTemporary
            strsql = " Insert Into InventoryTemp Select * From Inventory"
            CommandExecute
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A": ClearFrame: ClearItemBox
    LoadMISItemsList
End Sub
Private Sub CmdSaveMISDetails_Click()
   If Not DataItemValidation Then
        cmdSaveMISDetails_LostFocus
        Exit Sub
   End If
   MSSItemNo = MSSItemNo + 1
   CheckAvail
   SaveTemporary
   UpdateInventoryTemp
      'SaveStockCard
   LoadMISDetails
    
   ClearFrame
   cmdPrint.Enabled = True
   lvwMIS.Enabled = True
   
End Sub
Private Sub cmdCancel_Click()
DeleteTemporary: ClearBox: ClearItemBox: ClearFrame
    BoxState False: ButtonState False: cmdNew.Enabled = True: cmdSearch.Enabled = True
    lvwMIS.ListItems.Clear: cmdNew.SetFocus
End Sub
Private Sub cmdEdit_Click()
 MsgBox "EDIT"
End Sub
Private Sub cmdDelete_Click()
    mmsAdoCmd.CommandText = "Delete From MISDetails"
    Set mmsADORst = mmsAdoCmd.Execute
    MsgBox "delete database"
End Sub
Private Sub cmdSearch_Click()
   EncodeMode = "S": ClearBox: ClearItemBox: ClearFrame: DeleteTemporary: BoxState False
   LoadSearchList
   txtMISSearch.Text = "": cboMISSearch.SetFocus
End Sub
Private Sub cmdPrint_Click()
    If EncodeMode = "A" Or EncodeMode = "E" Then
      strsql = " Insert Into MISDetails Select * From MISDetailsTemp": CommandExecute
      strsql = "Delete From Inventory": CommandExecute
      strsql = " Insert Into Inventory Select * From InventoryTemp ": CommandExecute
        'WHERE Remark NOT like '" & "ASSET" & "'": CommandExecute
      'strsql = " Insert Into Asset Select * From AssetTemp": CommandExecute
      'strsql = " Insert Into StockCard Select * From StockCardTemp "
      'CommandExecute
    End If
    
ClearBox: ClearItemBox: ClearFrame
BoxState False: ButtonState False
cmdNew.Enabled = True: cmdSearch.Enabled = True

Load DataEnvironment1
If DataEnvironment1.rsCommand3.State <> 0 Then DataEnvironment1.rsCommand3.Close
  ReportMIS.Refresh
If ReportMIS.Visible = False Then ReportMIS.Show

Set mmsADORst = Nothing
End Sub
Private Sub cmdAddClose_Click()
    frameItems.Visible = False
    cmdAdd.SetFocus
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then

     FormMainMenu.Show: ClearBox: ClearItemBox: ClearFrame
     FormMainMenu.cmdExit.SetFocus: Unload Me
     
   Else
     Exit Sub
   End If
End Sub
Private Sub cmdNew_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       cmdCancel_Click
    End If
End Sub
Private Sub cmdSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       cmdCancel_Click
    End If
End Sub
Private Sub cmdPrint_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       cmdCancel_Click
    End If
End Sub
Private Sub cmdAddEquip_Click()
       'FormEquipments.Show
End Sub
Private Sub cmdAddItem_Click()
       FormItems.Show
End Sub
Private Sub cmdFixAsset_Click()
On Error GoTo LocalError

         strsql = "INSERT INTO AssetTemp    (  AssetId"
         strsql = strsql & "            , AssetName"
         strsql = strsql & "            , AssetGroup"
         strsql = strsql & "            , Unit"
         strsql = strsql & "            , Cost"
         strsql = strsql & "            , Avail"
         strsql = strsql & "            , AssetAmount"
         strsql = strsql & "            , ATFNum"
         strsql = strsql & "            , ATFDate"
         strsql = strsql & "            , AssetTag"
         strsql = strsql & "            , SerialNum"
         strsql = strsql & "            , PONum"
         strsql = strsql & "            , Status"
         strsql = strsql & "            , InPosition"
         strsql = strsql & "            , BorrowDate"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & ItemInvID
         strsql = strsql & ", '" & Replace$(txtAssetName.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetGroup.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetAvail.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetAmount.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetATF.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtATFDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetTag.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetSerial.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetPO.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$("IN", "'", "''") & "'"
         strsql = strsql & ", '" & Replace$("WAREHOUSE", "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtATFDate.Text, "'", "''") & "'"
         strsql = strsql & ")"
         
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
         
         txtAssetGroup.Clear:   txtAssetSerial.Text = "": txtAssetPO.Text = ""
         txtAssetATF.Text = "": txtATFDate.Text = "__/__/____": txtAssetTag.Text = ""
         'MsgBox "DELETE IN INVENTORY"
         frameItemAsset.Visible = False
         'OptInventory.SetFocus
LocalError:
    Exit Sub
End Sub
'---------------------------------------------------------------------------------
'               T E X T B O X   C O N T R O L S
'---------------------------------------------------------------------------------
Private Sub TxtMISdate_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      'SendKeys_
   ElseIf KeyAscii = 13 Then
      txtMISRequest.SetFocus
   End If
End Sub
Private Sub txtMISDate_GotFocus()
   txtMISDate.SelLength = Len(txtMISDate.Text)
End Sub
Private Sub txtMISRequest_GotFocus()
  txtMISRequest.SelLength = Len(txtMISRequest.Text)
End Sub
Private Sub txtMISRequest_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtMISRemarks.SetFocus
   End If
   If Not IsDate(txtMISDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtMISDate.SetFocus
        txtMISDate = Format$(Now, "mm/dd/yyyy")
  End If
End Sub
Private Sub txtMISRemarks_GotFocus()
  txtMISRemarks.SelLength = Len(txtMISRemarks.Text)
End Sub
Private Sub txtMISRemarks_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      If txtMISRemarks.Text = "" Then
         txtMISRemarks.Text = "-"
      End If
      LoadMISItemsList
   End If
End Sub
'---------------
Private Sub txtMISItem_DblClick()
   ClearFrame
   LoadMISItemsList
End Sub
Private Sub txtMISQty_GotFocus()
 txtMISQty.SelLength = Len(txtMISQty.Text)
End Sub
Private Sub txtMISQty_Change()
  TxtVal = txtMISQty.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtMISQty.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtMISQty_KeyPress(KeyAscii As Integer)
On Error GoTo LocalError
   If KeyAscii = 13 Then
     CheckAvail
   End If
LocalError:
    Exit Sub
End Sub
Private Sub CheckAvail()
       If CDbl(txtMISAvail.Text) < CDbl(txtMISQty.Text) Then
          MsgBox "Not enough Stock", vbExclamation, "Number of Stock"
          txtMISQty.Text = "": txtMISQty.SetFocus
       Else
          ItemCost = txtMISCost.Text
          ItemAmount = Format((txtMISQty.Text * ItemCost), "Standard")
          ItemBal = CDbl(txtMISAvail.Text) - CDbl(txtMISQty.Text)
          InvAmount = Format((ItemBal * ItemCost), "Standard")
          txtMISAmount.Text = ItemAmount
          txtMISDept.SetFocus
       End If
End Sub
' ------------------   C H A R G I N G --------------------------
'-------------------------------------
Private Sub txtMISDept_GotFocus()
    strsql = "SELECT Department FROM Department ORDER BY Department"
    CommandExecute
    With mmsADORst
    txtMISDept.Clear
    Do While Not .EOF
        txtMISDept.AddItem ![Department]: .MoveNext
    Loop
    txtMISDept.ListIndex = 0
    End With
End Sub
Private Sub txtMISDept_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtMISDept.Text = "ADMIN" Then
      txtMISArea.Text = txtMISDept
   End If
   txtMISAsset.SetFocus
End If
End Sub
Private Sub txtMISAsset_GotFocus()
    strsql = "SELECT Tag FROM AssetTag ORDER BY Tag"
    CommandExecute
    With mmsADORst
    txtMISAsset.Clear
    Do While Not .EOF
        txtMISAsset.AddItem ![Tag]: .MoveNext
    Loop
    txtMISAsset.ListIndex = 0
    End With
End Sub
Private Sub txtMISAsset_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMISTransact.SetFocus
End If
End Sub
Private Sub txtMISTransact_GotFocus()
    strsql = "SELECT * FROM Transact ORDER BY TransactID"
    CommandExecute
    With mmsADORst
    txtMISTransact.Clear
    Do While Not .EOF
        txtMISTransact.AddItem ![TransactName]: .MoveNext
    Loop
    txtMISTransact.ListIndex = 0
    End With
End Sub
Private Sub txtMISTransact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMISSeason.SetFocus
End If
End Sub
Private Sub txtMISSeason_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMISArea.SetFocus
End If
End Sub
Private Sub txtMISArea_GotFocus()
    strsql = "SELECT * FROM Area ORDER BY AreaID"
    CommandExecute
    With mmsADORst
    txtMISArea.Clear
    Do While Not .EOF
        txtMISArea.AddItem ![AreaName]: .MoveNext
    Loop
    txtMISArea.ListIndex = 0
    End With
End Sub
Private Sub txtMISArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMISBlock.SetFocus
End If
End Sub
Private Sub txtMISBlock_GotFocus()
    strsql = "SELECT * FROM Block ORDER BY BlockID"
    CommandExecute
    With mmsADORst
    txtMISBlock.Clear
    Do While Not .EOF
        txtMISBlock.AddItem ![BlockNum]: .MoveNext
    Loop
    txtMISBlock.ListIndex = 0
    End With
End Sub
Private Sub txtMISBlock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtMISEquip.SetFocus
End If
End Sub
Private Sub txtMISEquip_GotFocus()
    strsql = "SELECT * FROM Equipment ORDER BY EquipID"
    CommandExecute
    With mmsADORst
    txtMISEquip.Clear
    Do While Not .EOF
        txtMISEquip.AddItem ![EquipName]: .MoveNext
    Loop
    txtMISEquip.ListIndex = 0
    End With
End Sub
Private Sub txtMISEquip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdSaveMISDetails.SetFocus
End If
End Sub

'------------------------------------------------------------------------------------
'                               A S S E T S
'-----------------------------------------------------------------------------------
Private Sub txtAssetGroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtAssetSerial.SetFocus
End If
End Sub
Private Sub txtAssetGroup_Click()
   txtAssetSerial.SetFocus
End Sub
Private Sub txtAssetSerial_GotFocus()
   txtAssetSerial.Text = "S/N:"
   txtAssetSerial.SelStart = 7
End Sub
Private Sub txtAssetSerial_KeyPress(KeyAscii As Integer)
KeyAscii = ConvertUpper(KeyAscii)
If KeyAscii = 13 Then
   txtAssetPO.SetFocus
End If
End Sub
Private Sub txtAssetPO_GotFocus()
   txtAssetPO.Text = "PO #:"
   txtAssetPO.SelStart = 7
End Sub
Private Sub txtAssetPO_KeyPress(KeyAscii As Integer)
KeyAscii = ConvertUpper(KeyAscii)
If KeyAscii = 13 Then
   txtAssetATF.SetFocus
End If
End Sub
Private Sub txtAssetATF_GotFocus()
   txtAssetATF.Text = "ATF #:"
   txtAssetATF.SelStart = 7
End Sub
Private Sub txtAssetATF_KeyPress(KeyAscii As Integer)
KeyAscii = ConvertUpper(KeyAscii)
If KeyAscii = 13 Then
   txtATFDate.SetFocus
   txtATFDate = Format$(Now, "mm/dd/yyyy")
End If
End Sub
Private Sub txtATFDate_GotFocus()
    txtATFDate = Format$(Now, "mm/dd/yyyy")
End Sub
Private Sub txtATFDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   GetAssetTag
   txtAssetTag.SetFocus
End If
End Sub
Private Sub txtAssetTag_KeyPress(KeyAscii As Integer)
KeyAscii = ConvertUpper(KeyAscii)
If KeyAscii = 13 Then
   cmdFixAsset.SetFocus
End If
End Sub
Private Sub GetAssetTag()
On Error GoTo LocalError
  'AssetID = Format$(GetNextAssetID, "0000")
  AssetGroup = Left(txtAssetGroup.Text, 3)
  AssetATF = Right(txtAssetATF.Text, Len(txtAssetATF.Text) - 6)
  txtAssetTag.Text = AssetGroup & "3-" & AssetATF
LocalError:
    Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                      S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboMISSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtMISSearch.SetFocus
   End If
   If KeyAscii < 255 Then
      'SendKeys_
   End If
End Sub
Private Sub cboMISSearch_GotFocus()
  cboMISSearch.Text = "NUMBER"
End Sub
Private Sub txtMISSearch_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     lvwMISSearch.SetFocus
  End If
End Sub
Private Sub txtMISSearch_Change()
Dim MISSearchLI       As ListItem

On Error GoTo LocalError
    If cboMISSearch.Text = "NUMBER" Then
       strsql = "Select * from MISDetails where MISID like '" & txtMISSearch.Text & "%'" & "Order by MISNum"
    ElseIf cboMISSearch.Text = "DATE" Then
       strsql = "Select * from MISDetails where MISDate like '" & txtMISSearch.Text & "%'" & "Order by MISNum"
    ElseIf cboMISSearch.Text = "ITEM" Then
       strsql = "Select * from MISDetails where MISItem like '" & txtMISSearch.Text & "%'" & "Order by MISNum"
    ElseIf cboMISSearch.Text = "APPLIED" Then
       strsql = "Select * from MISDetails where MISPhase like '" & txtMISSearch.Text & "%'" & "Order by MISNum"
    End If
    
      
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LoadMISSearch
    
LocalError:
    Exit Sub
End Sub

'---------------- S E A R C H  L I S T V I E W ------------
Private Sub lvwMISSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
   With Item
        SearchLoad
   End With
End Sub
Private Sub lvwMISSearch_DblClick()
   frameSearch.Visible = False
   BoxState False
   SaveTemporary
   LoadMISDetails
   cmdPrint.Enabled = True
End Sub
Private Sub lvwMISSearch_KeyPress(KeyAscii As Integer)
On Error GoTo LocalError
   If KeyAscii = 13 Then
       lvwMISSearch_DblClick
       MISResult = lvwMISSearch.SelectedItem.Text
   End If
LocalError:
    Exit Sub
End Sub
Private Sub lvwMISSearch_GotFocus()
On Error GoTo LocalError
    SearchLoad
LocalError:
    Exit Sub
End Sub
Private Sub SearchLoad()
      txtMISNum.Text = lvwMISSearch.SelectedItem.Text
      txtMISDate.Text = Format$(lvwMISSearch.SelectedItem.SubItems(1), "mm/dd/yyyy")
      txtMISRequest.Text = lvwMISSearch.SelectedItem.SubItems(2)
      txtMISLocation.Text = lvwMISSearch.SelectedItem.SubItems(3)
End Sub
Private Sub LoadSearchList()
   frameSearch.Visible = True
       frameSearch.Top = lvwMIS.Top
       frameSearch.Left = lvwMIS.Left
       frameSearch.Height = 7750
       frameSearch.Width = lvwMIS.Width
       lvwMISSearch.Height = frameSearch.Height - 900
   lvwMISItems.ColumnHeaders.Clear
    With lvwMISSearch
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NUMBER", .Width * 0.12
        .ColumnHeaders.Add , , "DATE", .Width * 0.1
        .ColumnHeaders.Add , , "REQUEST BY", .Width * 0.2
        .ColumnHeaders.Add , , "LOCATION", .Width * 0#
        .ColumnHeaders.Add , , "ITEM", .Width * 0.3
        .ColumnHeaders.Add , , "APPLIED TO", .Width * 0.25
    End With

    strsql = "SELECT DISTINCT MISNum, MISDate, MISLocation" _
           & "     , MISRequest, MISItem, MISPhase" _
           & "     FROM MISDetails ORDER BY MISNum "
    CommandExecute
    
    LoadMISSearch
End Sub
Private Sub LoadMISSearch()
Dim MISSearchLI   As ListItem
On Error GoTo LocalError

    lvwMISSearch.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MISSearchLI = lvwMISSearch.ListItems.Add(, , !MISNum & "")
            MISSearchLI.SubItems(1) = !MISDate & ""
            MISSearchLI.SubItems(2) = !MISRequest & ""
            MISSearchLI.SubItems(3) = !MISLocation & ""
            MISSearchLI.SubItems(4) = !MISItem & ""
            MISSearchLI.SubItems(5) = !MISPhase & ""
            .MoveNext
        Loop
    End With
     
    Set MISSearchLI = Nothing
    Set mmsADORst = Nothing
    
LocalError:
    Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                      I T E M S   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtItemSearch_Change()
    strsql = "Select * from InventoryTemp where ItemName like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemName"
    
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwMISItems.ListItems.Clear
     LoadItems

    Set mmsADORst = Nothing
End Sub
Private Sub txtItemSearch_Click()
    txtItemSearch.Text = ""
End Sub
Private Sub txtItemSearch_GotFocus()
    txtItemSearch.Text = ""
   txtItemSearch.SelLength = Len(txtItemSearch.Text)
End Sub
Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
       cmdAddItem_Click
   End If
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      lvwMISItems.SetFocus
   End If
End Sub
Private Sub lvwMISItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      ItemGroup = lvwMISItems.SelectedItem.Text
      ItemCode = lvwMISItems.SelectedItem.SubItems(1)
      LoadItemSearch
    End With
End Sub
Private Sub lvwMiSItems_DblClick()
    frameItems.Visible = False:
    frameItemAdd.Top = lvwMIS.Top - 100: frameItemAdd.Left = 7500: frameItemAdd.Height = 7550: frameItemAdd.Width = 10500
    ItemBoxState True
    frameItemAdd.Visible = True: txtMISTransact_GotFocus: txtMISArea_GotFocus: txtMISBlock_GotFocus: txtMISEquip_GotFocus: txtMISAsset_GotFocus
    txtMISQty.SetFocus:
End Sub
Private Sub lvwMiSItems_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
      cmdAddItem_Click
   End If
End Sub
Private Sub lvwMiSItems_KeyPress(KeyAscii As Integer)
   lvwMiSItems_DblClick
End Sub
Private Sub lvwMiSItems_GotFocus()
On Error GoTo LocalError
      ItemGroup = lvwMISItems.SelectedItem.Text
      ItemCode = lvwMISItems.SelectedItem.SubItems(1)
      LoadItemSearch
LocalError:
    Exit Sub
End Sub
Private Sub LoadItemSearch()
      txtMISGroup.Text = ItemGroup
      txtMISItem.Text = lvwMISItems.SelectedItem.SubItems(2)
      txtMISAvail.Text = lvwMISItems.SelectedItem.SubItems(3)
      txtMISUnit.Text = lvwMISItems.SelectedItem.SubItems(4)
      txtMISCost.Text = lvwMISItems.SelectedItem.SubItems(5)
      txtMISAmount.Text = Format((CDbl(txtMISAvail.Text) * CDbl(txtMISCost.Text)), "Standard")
      ItemInvID = lvwMISItems.SelectedItem.SubItems(8)
End Sub
Private Sub LoadMISItemsList()
     If Not DataValidation Then
       Exit Sub
     Else
       
     End If
    
    BoxState False
    frameItems.Top = 750: frameItems.Left = lvwMIS.Left: frameItems.Height = 8750: frameItems.Width = lvwMIS.Width
    lvwMISItems.Top = 950: lvwMISItems.Left = 150:  lvwMISItems.Height = frameItems.Height - 1100
    frameItems.Visible = True: txtItemSearch.SetFocus
    
     lvwMISItems.ColumnHeaders.Clear
     With lvwMISItems
        .ColumnHeaders.Add , , "GROUP", .Width * 0.1
        .ColumnHeaders.Add , , "CODE", .Width * 0#
        .ColumnHeaders.Add , , "ITEM NAME", .Width * 0.25
        .ColumnHeaders.Add , , "STOCK", .Width * 0.1
        .ColumnHeaders.Add , , " ", .Width * 0.07
        .ColumnHeaders.Add , , "COST", .Width * 0.1
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.13
        .ColumnHeaders.Add , , "DETAILS", .Width * 0.2
        .ColumnHeaders.Add , , "INVID", .Width * 0#
     End With
    lvwMISItems.ColumnHeaders.Item(4).Alignment = lvwColumnRight: lvwMISItems.ColumnHeaders.Item(6).Alignment = lvwColumnRight
    lvwMISItems.ColumnHeaders.Item(7).Alignment = lvwColumnRight
    
    strsql = "SELECT * From InventoryTemp Order By ItemGroup, ItemName": CommandExecute
    lvwMISItems.ListItems.Clear
       LoadItems
    Set mmsADORst = Nothing
    
LocalError:
    Exit Sub
End Sub
Private Sub LoadItems()
    With mmsADORst
        Do Until .EOF
            Set LI = lvwMISItems.ListItems.Add(, , !ItemGroup & "")
                LI.SubItems(1) = !ItemCode & ""
                LI.SubItems(2) = !ItemName & ""
                LI.SubItems(3) = !AvailStock & ""
                LI.SubItems(4) = !Unit & ""
                LI.SubItems(5) = !Cost & ""
                LI.SubItems(6) = Format$(!ItemAmount, "#,###.#0") & ""
                LI.SubItems(7) = !Remarks & ""
                LI.SubItems(8) = !invid & ""
            .MoveNext
        Loop
     End With
End Sub
'--------------------------------------------------------------------------
'                     L I S T V I E W
'--------------------------------------------------------------------------
Private Sub lvwMIS_DblClick()
   If lvwMIS.SelectedItem Is Nothing Then
   Else
        If MsgBox("Delete this entry?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
            DeleteEntry
        Else
            Exit Sub
         End If
   End If
End Sub
Private Sub DeleteEntry()
    ItemNoDel = lvwMIS.SelectedItem
    ItemCode = lvwMIS.SelectedItem.SubItems(2)
    ItemQtyDel = lvwMIS.SelectedItem.SubItems(4)
    ItemInvID = lvwMIS.SelectedItem.SubItems(9)

        strsql = "Delete From MISDetailsTemp where MISItemNo like " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
    
        strsql = "Update MISDetailsTemp SET MISItemNo = MISItemNo - 1"
        strsql = strsql & " where MISItemNo > " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
        '-------------------
        'strsql = "Delete From StockCardTemp where StockID like " & ItemNoDel & ""
        'mmsAdoCmd.CommandText = strsql
        'Set mmsADORst = mmsAdoCmd.Execute
        '---------------------------
        strsql = "Select * From InventoryTemp where InvID like '" & ItemInvID & "'"
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
        ItemBal = CDbl(mmsADORst.Fields("AvailStock")) + CDbl(ItemQtyDel)
        ItemCost = mmsADORst.Fields("Cost")
        InvAmount = Format$(ItemBal * ItemCost, "#,###.#0")
        UpdateInventoryTemp
        
        MSSItemNo = MSSItemNo - 1
          
       LoadMISDetails
    
End Sub
Private Sub SetlvwMIS()
lvwMIS.ColumnHeaders.Clear
    With lvwMIS
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " # ", .Width * 0.03
        .ColumnHeaders.Add , , "GROUP", .Width * 0.08
        .ColumnHeaders.Add , , "CODE", .Width * 0#
        .ColumnHeaders.Add , , "ITEM", .Width * 0.3
        .ColumnHeaders.Add , , "QTY", Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.05
        .ColumnHeaders.Add , , "U/C", Width * 0.1
        .ColumnHeaders.Add , , "AMOUNT", Width * 0.13
        .ColumnHeaders.Add , , " ", Width * 0.48
        .ColumnHeaders.Add , , "ID ", Width * 0#
    End With
  lvwMIS.ColumnHeaders.Item(5).Alignment = lvwColumnRight: lvwMIS.ColumnHeaders.Item(7).Alignment = lvwColumnRight
  lvwMIS.ColumnHeaders.Item(8).Alignment = lvwColumnRight
End Sub
'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "Select * from MISDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub LoadMISDetails()
Dim MISLI             As ListItem
On Error GoTo LocalError

      GetMISTotal
       txtMISTotal.Text = Format(MISTotal, "Standard")
      InsertMISTotal
    
    strsql = "SELECT * from MISDetailsTemp ORDER BY MISItemNo "
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwMIS.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set MISLI = lvwMIS.ListItems.Add(, , !MISItemNo & "")
            MISLI.SubItems(1) = !MISGroup & ""
            MISLI.SubItems(2) = !MISCode & ""
            MISLI.SubItems(3) = !MISItem & ""
            MISLI.SubItems(4) = !MISQty & ""
            MISLI.SubItems(5) = !MISUnit & ""
            MISLI.SubItems(6) = !MISCost & ""
            MISLI.SubItems(7) = !MISAmount & ""
            MISLI.SubItems(8) = !MISInfoDetails & ""
            MISLI.SubItems(9) = !MISInvID & ""
            .MoveNext
        Loop
      End With
      
    cmdAdd.SetFocus
LocalError:
    Exit Sub
End Sub
Private Sub GetMISTotal()
        strsql = " SELECT SUM(MISAmount) as SubTotal FROM MISDetailsTemp "
        CommandExecute
        MISTotal = mmsADORst.Fields!Subtotal
        txtMISTotal.Text = Format(MISTotal, "Standard")
End Sub
Private Sub InsertMISTotal()
On Error GoTo LocalError
        strsql = "Update MISDetailsTemp SET MISTotal = '" & txtMISTotal.Text & "'"
        strsql = strsql & " where MISNum like '" & txtMISNum.Text & "'"
        CommandExecute
LocalError:
    Exit Sub
End Sub
Private Sub SaveTemporary()
On Error GoTo LocalError
    
     
     If EncodeMode = "A" Then
         CheckAvail
         If txtMISDept.Text = "ADMIN" Then
            txtMISArea.Text = txtMISDept.Text
         End If
  MISDetails = txtMISTransact.Text & " - " & txtMISSeason.Text & " - " & txtMISArea.Text & " - " & txtMISBlock.Text & " - " & txtMISEquip.Text
         
         strsql = "INSERT INTO MISDetailsTemp    (  MISId"
         strsql = strsql & "            , MISItemNo"
         strsql = strsql & "            , MISNum"
         strsql = strsql & "            , MISDate"
         strsql = strsql & "            , MISLocation"
         strsql = strsql & "            , MISRequest"
         strsql = strsql & "            , MISRemark"
         strsql = strsql & "            , MISGroup"
         strsql = strsql & "            , MISDept"
         strsql = strsql & "            , MISSeason"
         strsql = strsql & "            , MISTransact"
         strsql = strsql & "            , MISPhase"
         strsql = strsql & "            , MISBlock"
         strsql = strsql & "            , MISInfoDetails"
         strsql = strsql & "            , MISEquip"
         strsql = strsql & "            , MISItem"
         strsql = strsql & "            , MISCode"
         strsql = strsql & "            , MISQty"
         strsql = strsql & "            , MISUnit"
         strsql = strsql & "            , MISCost"
         strsql = strsql & "            , MISAmount"
         strsql = strsql & "            , MISInvId"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & MISId
         strsql = strsql & ", '" & MSSItemNo & "'"
         strsql = strsql & ", '" & Replace$(txtMISNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISLocation.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISRequest.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISRemarks.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISGroup.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISDept.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISSeason.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISTransact.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISArea.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISBlock.Text, "'", "''") & "'"
         strsql = strsql & ", '" & MISDetails & "'"
         strsql = strsql & ", '" & Replace$(txtMISEquip.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtMISQty.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemAmount & "'"
         strsql = strsql & ", '" & ItemInvID & "'"
         strsql = strsql & ")"
         
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
       
       If txtMISAsset.Text = "ASSET" Then
          frameItemAsset.Visible = True
          LoadAsset
          txtAssetGroup.SetFocus
       End If
           
   End If
   
  If EncodeMode = "S" Then
     strsql = " Insert Into MISDetailsTemp Select * From MISDetails Where MISNum like '" & txtMISNum.Text & "' Order By MISItemNo "
     mmsAdoCmd.CommandText = strsql
     Set mmsADORst = mmsAdoCmd.Execute
     cmdCancel.Enabled = True
  End If
   
LocalError:
    Exit Sub
End Sub
Private Sub LoadAsset()
    txtAssetName.Text = txtMISItem.Text: txtAssetAvail.Text = txtMISQty.Text: txtAssetUnit.Text = txtMISUnit.Text
    txtAssetCost.Text = txtMISCost.Text: txtAssetAmount.Text = txtMISAmount.Text: txtAssetGroup.SetFocus
End Sub
Private Sub SaveStockCard()
On Error GoTo LocalError
         
         strsql = "INSERT INTO StockCardTemp (  StockID"
         strsql = strsql & "            , StockDate"
         strsql = strsql & "            , StockCode"
         strsql = strsql & "            , StockItem"
         strsql = strsql & "            , StockCategory"
         strsql = strsql & "            , StockReorder"
         strsql = strsql & "            , StockNum"
         strsql = strsql & "            , StockPhase"
         strsql = strsql & "            , StockUnit"
         strsql = strsql & "            , StockCost"
         strsql = strsql & "            , StockAmount"
         strsql = strsql & "            , StockIn"
         strsql = strsql & "            , StockOut"
         strsql = strsql & "            , StockBalance"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & MISId
         strsql = strsql & ", '" & Replace$(txtMISDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtMISItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemGroup & "'"
             strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISArea.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemAmount & "'"
          strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMISQty.Text, "'", "''") & "'"
             strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
         strsql = strsql & ")"
         
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
  
LocalError:
    Exit Sub
End Sub
Private Sub UpdateInventoryTemp()
On Error GoTo LocalError
        strsql = "Update InventoryTemp SET Cost = '" & ItemCost & "'"
        strsql = strsql & ", AvailStock = '" & ItemBal & "'"
        strsql = strsql & ", ItemAmount = '" & InvAmount & "'"
        'strsql = strsql & ", Remarks    = '" & txtMISNum.Text & "-" & txtMISDate.Text & "/" & txtMISAsset.Text & "'"
        strsql = strsql & " WHERE InvID like '" & ItemInvID & "'"
        CommandExecute
LocalError:
    Exit Sub
End Sub
Private Sub UpdateAssetTemp()
       If txtMISAsset.Text = "ASSET" Then
          frameItemAsset.Visible = True
          LoadAsset
          txtAssetGroup.SetFocus
       End If
End Sub
Private Sub CalculateTotal()
Dim i As Double
    MISTotal = 0
    For i = 1 To lvwMIS.ListItems.Count
       MISTotal = MISTotal + CCur(lvwMIS.ListItems(i).SubItems(4)) ' Ammount Column
    Next i
    txtMISTotal.Text = Format$(MISTotal, "#,###.#0")
End Sub
Private Sub DeleteTemporary()
On Error GoTo LocalError
    strsql = "Delete From MISDetailsTemp": CommandExecute
    strsql = "Delete From StockCardTemp": CommandExecute
    strsql = "Delete From InventoryTemp": CommandExecute
    strsql = "Delete From AssetTemp": CommandExecute
LocalError:
Exit Sub
End Sub
Private Function GetNextMISID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(MISID) AS MaxID FROM MISDetails"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextMISID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextMISID = 1
       Else
           GetNextMISID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtMISDate.Text = "" Then
        MsgBox "Fill-up MIS Date.", vbExclamation, "MIS Date Required"
        txtMISDate.SetFocus
        Exit Function
    End If
    If txtMISRequest.Text = "" Then
        MsgBox "Fill-up Person Received.", vbExclamation, "Requested Required"
        txtMISRequest.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function DataItemValidation() As Boolean
On Error GoTo LocalError
  DataItemValidation = False
    If txtMISQty.Text = "" Or txtMISQty.Text = "0" Then
        MsgBox "No Quantity of Item.", vbExclamation, "Quantity Required"
        txtMISQty.SetFocus
        Exit Function
    End If
    DataItemValidation = True
LocalError:
    'MsgBox "CHECK DATA"
    Exit Function
End Function
Private Sub SendKeys_()
    'SendKeys "{left}"
    'SendKeys "{del}"
End Sub
Private Sub ClearBox()
    txtMISNum.Text = ""
    txtMISDate.Text = "__/__/____"
    txtMISLocation.Text = ""
    txtMISRequest.Text = ""
    txtMISRemarks.Text = ""
    txtMISTotal.Text = ""
    lvwMIS.ListItems.Clear
End Sub
Private Sub ClearFrame()
    frameItemAdd.Visible = False
    frameSearch.Visible = False
    frameItems.Visible = False
    frameItemAsset.Visible = False
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtMISDate.Enabled = boxEnabled
    txtMISLocation.Enabled = boxEnabled
    txtMISRequest.Enabled = boxEnabled
    txtMISRemarks.Enabled = boxEnabled
End Sub
Private Sub ClearItemBox()
    ItemCode = ""
    ItemGroup = ""
    txtMISDept.Clear
    txtMISTransact.Clear
    txtMISArea.Clear
    txtMISBlock.Clear
    txtMISEquip.Clear
    txtMISItem.Text = ""
    txtMISQty.Text = ""
    txtMISUnit.Text = ""
    txtMISCost.Text = ""
    txtMISTotal.Text = ""
End Sub
Private Sub ItemBoxState(boxEnabled As Boolean)
    txtMISGroup.Enabled = boxEnabled
    txtMISDept.Enabled = boxEnabled
    txtMISTransact.Enabled = boxEnabled
    txtMISArea.Enabled = boxEnabled
    txtMISBlock.Enabled = boxEnabled
    txtMISEquip.Enabled = boxEnabled
    txtMISItem.Enabled = boxEnabled
    txtMISQty.Enabled = boxEnabled
    txtMISUnit.Enabled = boxEnabled
    txtMISCost.Enabled = boxEnabled
    txtMISAvail.Enabled = boxEnabled
    txtMISAmount.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwMIS.Enabled = buttonEnabled
    cmdNew.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdSearch.Enabled = buttonEnabled
    cmdPrint.Enabled = buttonEnabled
    'cmdEdit.Enabled = buttonEnabled
    'cmdDelete.Enabled = buttonEnabled
End Sub
'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HC0FFC0
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdEdit_GotFocus()
   cmdEdit.BackColor = &HC0FFC0
End Sub
Private Sub cmdEdit_LostFocus()
   cmdEdit.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0FFC0
End Sub
Private Sub cmdSaveMISDetails_GotFocus()
   cmdSaveMISDetails.BackColor = &HC0FFFF
End Sub
Private Sub cmdSaveMISDetails_LostFocus()
   cmdSaveMISDetails.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HC0FFC0
End Sub
Private Sub cmdSearch_LostFocus()
   cmdSearch.BackColor = &H8000000F
End Sub
Private Sub cmdPrint_GotFocus()
   cmdPrint.BackColor = &HC0FFC0
End Sub
Private Sub cmdPrint_LostFocus()
   cmdPrint.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub cmdEditSave_GotFocus()
   'cmdEditSave.BackColor = &HC0FFC0
End Sub
Private Sub cmdEditSave_LostFocus()
   'cmdEditSave.BackColor = &H8000000F
End Sub


