VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTS "
   ClientHeight    =   10800
   ClientLeft      =   2640
   ClientTop       =   375
   ClientWidth     =   16905
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framePODetails 
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
      ForeColor       =   &H00004000&
      Height          =   3900
      Left            =   1200
      TabIndex        =   39
      Top             =   3600
      Visible         =   0   'False
      Width           =   7045
      Begin VB.TextBox txtPOGroup 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   400
         Width           =   4740
      End
      Begin VB.TextBox txtPOCost 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   43
         Top             =   1680
         Width           =   4770
      End
      Begin VB.TextBox txtPOItem 
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
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Top             =   1050
         Width           =   4740
      End
      Begin VB.CommandButton cmdSavePODetails 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3120
         Width           =   6360
      End
      Begin VB.TextBox txtPOAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Top             =   2280
         Width           =   4770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "COST "
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
         Index           =   13
         Left            =   480
         TabIndex        =   47
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "QTY"
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
         Left            =   480
         TabIndex        =   46
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "AMOUNT "
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
         Index           =   14
         Left            =   480
         TabIndex        =   45
         Top             =   1920
         Width           =   1080
      End
   End
   Begin VB.Frame FrameMIS 
      BackColor       =   &H00C0C0C0&
      Height          =   1245
      Left            =   12120
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   4125
      Begin VB.OptionButton OptMISCharge 
         BackColor       =   &H00FFFFC0&
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton OptMISDept 
         BackColor       =   &H00FFFFC0&
         Caption         =   "DEPT."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptMISCCC 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CCC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMISNumber 
         BackColor       =   &H00FFFFC0&
         Caption         =   "NUM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton OptMISBlk 
         BackColor       =   &H00FFFFC0&
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptMISSpv 
         BackColor       =   &H00FFFFC0&
         Caption         =   "RECEIVED"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptBlk 
         BackColor       =   &H00FFFFC0&
         Caption         =   "BLK_"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2160
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "MIS REPORTS"
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
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.Frame FrameMRR 
      BackColor       =   &H00C0C0C0&
      Height          =   1245
      Left            =   10200
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
      Begin VB.OptionButton OptMRRNumber 
         BackColor       =   &H00FFFFC0&
         Caption         =   "NUMBER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton OptMRRSupplier 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "MRR REPORTS"
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
         Left            =   120
         TabIndex        =   24
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   2500
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   2500
   End
   Begin VB.Frame frameItemOpt 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2370
      Left            =   13680
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CommandButton cmdConvert 
         Caption         =   "CONVERT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   2500
      End
      Begin VB.CommandButton cmdClearInv 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1500
         Width           =   2500
      End
      Begin VB.CommandButton cmdAsset 
         Caption         =   "ASSET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   900
         Width           =   2500
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
         Height          =   240
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1400
      Left            =   5880
      TabIndex        =   12
      Top             =   0
      Width           =   11295
      Begin VB.OptionButton OptAssets 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FIXED ASSETS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Frame FramePO 
         BackColor       =   &H00C0C0C0&
         Height          =   1365
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   5760
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "Report.frx":27337
            Left            =   1560
            List            =   "Report.frx":2734D
            TabIndex        =   36
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton OptSummary 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SUMMARY"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   270
            Left            =   3960
            TabIndex        =   35
            Top             =   300
            Width           =   1395
         End
         Begin VB.ComboBox cboPO 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "Report.frx":27383
            Left            =   1560
            List            =   "Report.frx":27399
            TabIndex        =   34
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptItem 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ITEMIZED"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   270
            Left            =   3960
            TabIndex        =   20
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "TRANSFER"
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
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "P.O."
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
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.OptionButton OptFuel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FUEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton OptInventory 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INVENTORY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1400
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtYear 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "Report.frx":273CF
         Left            =   200
         List            =   "Report.frx":273F7
         TabIndex        =   0
         Top             =   650
         Width           =   2300
      End
      Begin MSMask.MaskEdBox txtDRStart 
         Height          =   420
         Left            =   2600
         TabIndex        =   1
         Top             =   650
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDREnding 
         Height          =   420
         Left            =   4100
         TabIndex        =   2
         Top             =   650
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REPORT RANGE"
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
         Left            =   200
         TabIndex        =   11
         Top             =   240
         Width           =   1590
      End
   End
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   9435
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   16642
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   16935
   End
End
Attribute VB_Name = "FormReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn          As ADODB.Connection
Private mmsAdoCmd           As ADODB.Command
Private mmsADORst           As ADODB.Recordset
Private dbcommand           As ADODB.Command
Private strsql              As String

Dim SummaryLI, InventoryLI          As ListItem
Dim SummaryRow                      As Integer
 
Dim ItemClick, Unit, TxtVal, NumVal, ConRemark, AssetID, AssetGroup, AssetATF                 As String
Dim DeptItem, SupplierItem, MatItem, NumItem, TransactItem, PhaseItem, EquipItem, MRRItem     As String
Dim POOpt, Header, GroupItem, LocationItem          As String

Dim SummaryTotal, DeptPhaseSum, DeptGroupSum, NumSubTotal, ChargeTotal, VehicleTotal, PhaseGroupSum, QtyTotal  As Currency
Dim i, Cost, AveCost, ItemInvID, MRSStock, MRSAmount, MISSTock, MISAmount, PreStock, PreAmount, PreCost, ItemAmount  As Double

Public ReportTransact, ReportType, ReportTittle, MonthSummary, StartMonth, EndMonth As String


'-----------------------------------------------------------------------------------
'                           F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
    Load Me
    ConnectToDB
    lblComp.Caption = FormMainMenu.lblComp.Caption: GetCurrentDate
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub
'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "Select * from MRSDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
'---------------------------------------------------------------------------------
'                         C O N T R O L S   E V E N T S
'---------------------------------------------------------------------------------
Private Sub lvwSummary_DblClick()
    If cboPO.Text = "NUMBER" And OptSummary.Value = True Then
        If Not lvwSummary.SelectedItem = " " Then
            framePODetails.Visible = False
        End If
    Else
        MsgBox "Trial"
    End If
End Sub
Private Sub cboMonth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDRStart.SetFocus
   Else
      cboMonth.Text = ""
   End If
End Sub
Private Sub cboMonth_GotFocus()
 cboMonth.SelLength = Len(cboMonth.Text)
 GetCurrentDate
 ClearOptions
End Sub
Private Sub cboMonth_Click()
   lvwSummary.ListItems.Clear
   ReportRange
End Sub
Private Sub cboMonth_LostFocus()
   If (Month(Now)) = 1 Then
       txtYear.Text = ((Year(Now) - 1))
   Else
       txtYear.Text = Year(Now)
   End If
   lvwSummary.ListItems.Clear
   ReportRange
End Sub
Private Sub txtDRStart_GotFocus()
 txtDRStart.SelLength = Len(txtDRStart.Text)
End Sub
Private Sub txtDRStart_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
   ElseIf KeyAscii = 13 Then
     If Not IsDate(txtDRStart.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDRStart.SetFocus
        txtDRStart.Text = Format$(StartMonth, "mm/dd/yyyy")
     Else
        txtDREnding.SetFocus
     End If
   End If
End Sub
Private Sub txtDREnding_GotFocus()
 txtDREnding.SelLength = Len(txtDREnding.Text)
End Sub
Private Sub txtDREnding_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
   ElseIf KeyAscii = 13 Then
     If Not IsDate(txtDREnding.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDREnding.SetFocus
        txtDREnding.Text = Format$(EndMonth, "mm/dd/yyyy")
     Else
        'txtYear.SetFocus
     End If
   End If
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      'OptDRSale.SetFocus
    Else
       If IsNumeric(Chr(KeyAscii)) Then
       Else
       End If
    End If
   lvwSummary.ListItems.Clear
End Sub
Private Sub txtYear_LostFocus()
   lvwSummary.ListItems.Clear
   'ReportRange
End Sub
Private Sub txtYear_GotFocus()
   txtYear.SelStart = 0
   txtYear.SelLength = Len(txtYear.Text)
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub cmdPrint_Click()
    cmdPrint_LostFocus
    cmdExit.SetFocus
    ListViewPrint
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     Unload Me
 Else
     Exit Sub
 End If
End Sub
'---------------------------------------------------------------------------
'                         P O    O  P  T  I  O  N  S
'--------------------------------------------------------------------------
Private Sub POHead()
ReportTransact = "PO"
ReportTittle = "PURCHASE ORDER"
End Sub
' ------------------- S U M M A R Y -----------------------
Private Sub OptSummary_GotFocus()
On Error GoTo LocalError
OptSummary.FontBold = True: OptItem.FontBold = False
StartMonth = txtDRStart.Text: EndMonth = txtDREnding.Text: ReportType = cboPO.Text: POOpt = "SUMMARY": POHead: SetlvwPO

If cboPO.Text = "NUMBER" Then
  strsql = "SELECT Distinct PONum, PODate, POArea, POPrs, POSupplier, POTerms, POWork, POEquip, POTotal ,POStatus, POTr, POMrr From PODetails " _
   & " WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY PONum"
  CommandExecute
    lvwSummary.ForeColor = &H0&
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum & "")
              SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POTerms
              SummaryLI.SubItems(4) = !POPrs: SummaryLI.SubItems(5) = !POWork: SummaryLI.SubItems(6) = !POEquip
              SummaryLI.SubItems(7) = Format$(!POTotal, "#,###.#0"): SummaryLI.SubItems(8) = !POStatus
              SummaryLI.SubItems(9) = !POTr: SummaryLI.SubItems(10) = !POMrr
            .MoveNext
        Loop
    End With
POTotal
SummaryLI.SubItems(7) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(7).ForeColor = &HC0&
End If

If cboPO.Text = "SUPPLIER" Then
   SetlvwPO2
   strsql = "SELECT DISTINCT POSupplier From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
   CommandExecute
    With mmsADORst
     Do Until .EOF
        Set SummaryLI = lvwSummary.ListItems.Add(, , !POSupplier & "")
         GroupItem = !POSupplier: SummaryLI.Bold = True
         POSumTotal
         SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
         .MoveNext
     Loop
    End With
POTotal
SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
End If

If cboPO.Text = "AREA" Then
   strsql = "SELECT DISTINCT POArea From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
   CommandExecute
    With mmsADORst
     Do Until .EOF
        Set SummaryLI = lvwSummary.ListItems.Add(, , !POArea & "")
         GroupItem = !POArea: SummaryLI.Bold = True
         Set SummaryLI = lvwSummary.ListItems.Add(, , "SERVED")
         GetAreaServed
         Set SummaryLI = lvwSummary.ListItems.Add(, , "PENDING")
         GetAreaPending
         .MoveNext
         LVSpace
     Loop
    End With
POTotal
SummaryLI.SubItems(8) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
End If

If cboPO.Text = "DETAILS" Then
    SetlvwPODetails
    lvwSummary.ColumnHeaders.Item(2).Alignment = lvwColumnRight: lvwSummary.ColumnHeaders.Item(3).Alignment = lvwColumnRight
    lvwSummary.ColumnHeaders.Item(4).Alignment = lvwColumnRight: lvwSummary.ColumnHeaders.Item(5).Alignment = lvwColumnRight
    lvwSummary.ColumnHeaders.Item(6).Alignment = lvwColumnRight: lvwSummary.ColumnHeaders.Item(7).Alignment = lvwColumnRight
    lvwSummary.ColumnHeaders.Item(8).Alignment = lvwColumnRight
    
   strsql = "SELECT DISTINCT POWork From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
   CommandExecute
    With mmsADORst
     Do Until .EOF
        Set SummaryLI = lvwSummary.ListItems.Add(, , !POWork & "")
         GroupItem = !POWork: SummaryLI.Bold = True
         LocationItem = lvwSummary.ColumnHeaders.Item(2): GetWorkArea: SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0")
         LocationItem = lvwSummary.ColumnHeaders.Item(3): GetWorkArea: SummaryLI.SubItems(2) = Format$(SummaryTotal, "#,###.#0")
         LocationItem = lvwSummary.ColumnHeaders.Item(4): GetWorkArea: SummaryLI.SubItems(3) = Format$(SummaryTotal, "#,###.#0")
         LocationItem = lvwSummary.ColumnHeaders.Item(5): GetWorkArea: SummaryLI.SubItems(4) = Format$(SummaryTotal, "#,###.#0")
         LocationItem = lvwSummary.ColumnHeaders.Item(6): GetWorkArea: SummaryLI.SubItems(5) = Format$(SummaryTotal, "#,###.#0")
         LocationItem = lvwSummary.ColumnHeaders.Item(7): GetWorkArea: SummaryLI.SubItems(6) = Format$(SummaryTotal, "#,###.#0")
         GetWorkTotal
         SummaryLI.SubItems(7) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(7).ForeColor = &HC0&
         .MoveNext
     Loop
    End With
    Set SummaryLI = lvwSummary.ListItems.Add(, , "")
         LocationItem = lvwSummary.ColumnHeaders.Item(2): GetAreaTotal: SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
         LocationItem = lvwSummary.ColumnHeaders.Item(3): GetAreaTotal: SummaryLI.SubItems(2) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(2).ForeColor = &HC0&
         LocationItem = lvwSummary.ColumnHeaders.Item(4): GetAreaTotal: SummaryLI.SubItems(3) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(3).ForeColor = &HC0&
         LocationItem = lvwSummary.ColumnHeaders.Item(5): GetAreaTotal: SummaryLI.SubItems(4) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(4).ForeColor = &HC0&
         LocationItem = lvwSummary.ColumnHeaders.Item(6): GetAreaTotal: SummaryLI.SubItems(5) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(5).ForeColor = &HC0&
         LocationItem = lvwSummary.ColumnHeaders.Item(7): GetAreaTotal: SummaryLI.SubItems(6) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(6).ForeColor = &HC0&
GetAreaTotal
SummaryLI.SubItems(6) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(6).ForeColor = &HC0&
GetPOTotal
SummaryLI.SubItems(7) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(7).ForeColor = &HC0&
End If

If cboPO.Text = "VEHICLE" Then
   SetlvwPO2
   strsql = "SELECT DISTINCT POEquip From PODetails WHERE POEquip NOT LIKE '" & "-" & "' AND PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
   CommandExecute
    With mmsADORst
     Do Until .EOF
        Set SummaryLI = lvwSummary.ListItems.Add(, , !POEquip & "")
         GroupItem = !POEquip: SummaryLI.Bold = True
         POSumTotal
         SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
         .MoveNext
     Loop
    End With
POTotal
SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
End If

If cboPO.Text = "STATUS" Then
  MsgBox "BY AREA"
   SetlvwPO2
   strsql = "SELECT DISTINCT POStatus From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
   CommandExecute
    With mmsADORst
     Do Until .EOF
        Set SummaryLI = lvwSummary.ListItems.Add(, , !POStatus & "")
         GroupItem = !POStatus: SummaryLI.Bold = True
         POSumTotal
         SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
         .MoveNext
     Loop
    End With
POTotal
SummaryLI.SubItems(1) = Format$(SummaryTotal, "#,###.#0"): SummaryLI.ListSubItems(1).ForeColor = &HC0&
End If

LocalError:
    Exit Sub
End Sub
Private Sub GetAreaServed()
  strsql = "SELECT Distinct PONum, PODate, POArea, POPrs, POSupplier, POTerms, POWork, POEquip, POTotal ,POStatus From PODetails " _
   & " WHERE POArea like '" & GroupItem & "' AND POStatus like '" & "SERVED" & "' AND PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY PONum"
  CommandExecute
    lvwSummary.ForeColor = &H0&
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum & "")
              SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POTerms
              SummaryLI.SubItems(4) = !POArea: SummaryLI.SubItems(5) = !POPrs: SummaryLI.SubItems(6) = !POWork
              SummaryLI.SubItems(7) = !POEquip: SummaryLI.SubItems(8) = Format$(!POTotal, "#,###.#0"): SummaryLI.SubItems(9) = !POStatus
              'SummaryLI.ListSubItems(8).ForeColor = &HC0&:
            .MoveNext
        Loop
    End With
End Sub
Private Sub GetAreaPending()
  strsql = "SELECT Distinct PONum, PODate, POArea, POPrs, POSupplier, POTerms, POWork, POEquip, POTotal ,POStatus From PODetails " _
   & " WHERE POArea like '" & GroupItem & "' AND POStatus not like '" & "SERVED" & "' AND PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY PONum"
  CommandExecute
    lvwSummary.ForeColor = &H0&
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum & "")
              SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POTerms
              SummaryLI.SubItems(4) = !POArea: SummaryLI.SubItems(5) = !POPrs: SummaryLI.SubItems(6) = !POWork
              SummaryLI.SubItems(7) = !POEquip: SummaryLI.SubItems(8) = Format$(!POTotal, "#,###.#0"): SummaryLI.SubItems(9) = !POStatus
              'SummaryLI.ListSubItems(8).ForeColor = &HC0&:
            .MoveNext
        Loop
    End With
End Sub
Private Sub GetWorkArea()
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
       & " AND POWork like '" & GroupItem & "' AND POArea like '" & LocationItem & "'": CommandExecute
  SummaryTotal = mmsADORst.Fields!Subtotal
End Sub
Private Sub GetWorkTotal()
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
       & " AND POWork like '" & GroupItem & "'": CommandExecute
  SummaryTotal = mmsADORst.Fields!Subtotal
End Sub
Private Sub GetAreaTotal()
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
       & " AND POArea like '" & LocationItem & "'": CommandExecute
  SummaryTotal = mmsADORst.Fields!Subtotal
End Sub
Private Sub GetPOTotal()
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# ": CommandExecute
  SummaryTotal = mmsADORst.Fields!Subtotal
End Sub
Private Sub POSumTotal()
If cboPO.Text = "SUPPLIER" Then
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
       & " AND POSupplier like '" & GroupItem & "'": CommandExecute
End If

If cboPO.Text = "AREA" Then
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POArea like '" & GroupItem & "'"
End If

If cboPO.Text = "DETAILS" Then
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POWork like '" & GroupItem & "'"
End If

If cboPO.Text = "VEHICLE" Then
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POEquip like '" & GroupItem & "'"
End If

If cboPO.Text = "STATUS" Then
  strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POStatus like '" & GroupItem & "'"
End If

CommandExecute
SummaryTotal = mmsADORst.Fields!Subtotal
End Sub
' ------------------- I T E M I Z E D -----------------------
Private Sub OptItem_GotFocus()
On Error GoTo LocalError
OptItem.FontBold = True: OptSummary.FontBold = False: SetlvwPO1
StartMonth = txtDRStart.Text: EndMonth = txtDREnding.Text: ReportType = cboPO.Text: POOpt = "ITEM": POHead

If cboPO.Text = "NUMBER" Then
  strsql = "SELECT Distinct PONum, PODate, POSupplier From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY PONum": CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum & " (" & !PODate & ") - ")
             GroupItem = !PONum: POItemized
            .MoveNext
            LVSpace
        Loop
    End With
End If

If cboPO.Text = "SUPPLIER" Then
  strsql = "SELECT DISTINCT POSupplier From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY POSupplier": CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !POSupplier & "")
              GroupItem = !POSupplier: POItemized
              .MoveNext
              LVSpace
        Loop
    End With
End If


If cboPO.Text = "AREA" Then
  strsql = "SELECT DISTINCT POArea From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY POArea": CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !POArea & "")
              GroupItem = !POArea: POItemized
              .MoveNext
              LVSpace
        Loop
    End With
End If

If cboPO.Text = "DETAILS" Then
  strsql = "SELECT DISTINCT POWork From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY POWork": CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !POWork & "")
              GroupItem = !POWork: POItemized
              .MoveNext
              LVSpace
        Loop
    End With
End If


If cboPO.Text = "VEHICLE" Then
  strsql = "SELECT DISTINCT POEquip From PODetails WHERE POEquip NOT LIKE '" & "-" & "' AND PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY POEquip"
  CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !POEquip & "")
              GroupItem = !POEquip: POItemized
              .MoveNext
              LVSpace
        Loop
    End With
End If

If cboPO.Text = "STATUS" Then
  strsql = "SELECT DISTINCT POStatus From PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY POStatus": CommandExecute
    With mmsADORst
        Do Until .EOF
          Set SummaryLI = lvwSummary.ListItems.Add(, , !POStatus & "")
              GroupItem = !POStatus: POItemized
              .MoveNext
              LVSpace
        Loop
    End With
End If

POTotal
LocalError: Exit Sub
End Sub
' -------------------  P O  ITEMIZED -----------------------
Private Sub POItemized()
Dim Subtotal As Currency
On Error GoTo LocalError

 If cboPO.Text = "NUMBER" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND PONum like '" & GroupItem & "' ORDER BY POItemNo"
   CommandExecute
    With mmsADORst
     Do Until .EOF
      Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
        SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
        SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
        SummaryLI.SubItems(7) = !POAmount
        .MoveNext
     Loop
    End With
    
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND PONum like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
 End If
 
 
 
  If cboPO.Text = "SUPPLIER" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND POSupplier like '" & GroupItem & "' ORDER BY PONum"
    CommandExecute
        With mmsADORst
         Do Until .EOF
            Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
                SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
                SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
                SummaryLI.SubItems(7) = !POAmount
               .MoveNext
        Loop
       End With
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POSupplier like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
  End If
  
  
  If cboPO.Text = "AREA" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND POArea like '" & GroupItem & "' ORDER BY PONum"
    CommandExecute
        With mmsADORst
         Do Until .EOF
            Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
                SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
                SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
                SummaryLI.SubItems(7) = !POAmount
               .MoveNext
        Loop
       End With
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails " _
          & " WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POArea like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
  End If
  
  If cboPO.Text = "DETAILS" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#  " _
           & " AND POWork like '" & GroupItem & "' ORDER BY POArea, PONum"
    CommandExecute
        With mmsADORst
         Do Until .EOF
            Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
                SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
                SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
                SummaryLI.SubItems(7) = !POAmount: SummaryLI.SubItems(8) = !POArea
               .MoveNext
        Loop
       End With
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails " _
          & " WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POWork like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
  End If
  
 If cboPO.Text = "VEHICLE" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#  " _
           & " AND POEquip like '" & GroupItem & "' ORDER BY PONum"
    CommandExecute
        With mmsADORst
         Do Until .EOF
            Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
                SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
                SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
                SummaryLI.SubItems(7) = !POAmount
               .MoveNext
        Loop
       End With
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails " _
          & " WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POEquip like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
  End If
  
  If cboPO.Text = "STATUS" Then
    strsql = "SELECT * FROM PODetails WHERE PODate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#  " _
           & " AND POStatus like '" & GroupItem & "' ORDER BY PONum"
    CommandExecute
        With mmsADORst
         Do Until .EOF
            Set SummaryLI = lvwSummary.ListItems.Add(, , !PONum)
                SummaryLI.SubItems(1) = !PODate: SummaryLI.SubItems(2) = !POSupplier: SummaryLI.SubItems(3) = !POQty
                SummaryLI.SubItems(4) = !POUnit: SummaryLI.SubItems(5) = !POItem: SummaryLI.SubItems(6) = !POCost
                SummaryLI.SubItems(7) = !POAmount
               .MoveNext
        Loop
       End With
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetails " _
          & " WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND POStatus like '" & GroupItem & "'"
         CommandExecute
         Subtotal = mmsADORst.Fields!Subtotal
         SummaryLI.SubItems(8) = Format$(Subtotal, "#,###.#0"): SummaryLI.ListSubItems(8).ForeColor = &HC0&
  End If
 
 
 
LocalError:
    Exit Sub
End Sub
Private Sub POTotal()
On Error GoTo LocalError

If Not cboPO.Text = "VEHICLE" Then
  strsql = " SELECT SUM(POAmount) as SummaryTotal FROM PODetails  WHERE PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# ": CommandExecute
Else
  strsql = " SELECT SUM(POAmount) as SummaryTotal FROM PODetails  WHERE POEquip NOT LIKE '" & "-" & "' AND PODate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
  CommandExecute
End If
SummaryTotal = mmsADORst.Fields!SummaryTotal

LVSpace2
LocalError: Exit Sub
End Sub
Private Sub BoldRed()
    SummaryLI.ListSubItems(6).Bold = True: SummaryLI.ListSubItems(6).ForeColor = &HC0&
    SummaryLI.ListSubItems(7).Bold = True: SummaryLI.ListSubItems(7).ForeColor = &HC0&
End Sub

'---------------------------------------------------------------------------
Private Sub LVSpace()
    Set SummaryLI = lvwSummary.ListItems.Add(, , " ")
End Sub
Private Sub LVSpace2()
LVSpace: LVSpace: LVSpace
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
Private Sub CommandExecute()
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
End Sub


