VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormTransfer 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSMITTAL"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10635
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameItemAdd 
      BackColor       =   &H00C0E0FF&
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
      Height          =   4620
      Left            =   4800
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   7045
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3600
         Width           =   3300
      End
      Begin VB.TextBox txtTrItemNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   720
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtTrUnit 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   720
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1410
      End
      Begin VB.TextBox txtTrCost 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         TabIndex        =   38
         Top             =   1800
         Width           =   4650
      End
      Begin VB.TextBox txtTrItem 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         Top             =   360
         Width           =   6300
      End
      Begin VB.TextBox txtTrQty 
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
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   36
         Top             =   1080
         Width           =   3210
      End
      Begin VB.CommandButton cmdSaveTRDetails 
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
         Left            =   3550
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3600
         Width           =   3300
      End
      Begin VB.TextBox txtTrAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
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
         TabIndex        =   34
         Top             =   2400
         Width           =   4650
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   0
         TabIndex        =   45
         Top             =   3360
         Width           =   7030
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   42
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   41
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Left            =   360
         TabIndex        =   40
         Top             =   2640
         Width           =   1080
      End
   End
   Begin VB.Frame frameOpt 
      BackColor       =   &H00C0E0FF&
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   46
      Top             =   6480
      Visible         =   0   'False
      Width           =   7170
      Begin VB.CommandButton cmdPO 
         Caption         =   "PURCHASE ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   3300
      End
      Begin VB.CommandButton cmdCV 
         Caption         =   "CASH VALE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   250
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label lblOpt 
         Caption         =   "Label2"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame framePO 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
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
      Height          =   1035
      Left            =   2160
      TabIndex        =   26
      Top             =   7920
      Visible         =   0   'False
      Width           =   12848
      Begin VB.TextBox txtPOSearch 
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
         Left            =   240
         MaxLength       =   50
         TabIndex        =   27
         Top             =   120
         Width           =   12405
      End
      Begin MSComctlLib.ListView lvwPO 
         Height          =   4230
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   12405
         _ExtentX        =   21881
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
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
   Begin VB.TextBox txtTrTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   850
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      Top             =   9600
      Width           =   3540
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   0
      TabIndex        =   22
      Top             =   9600
      Width           =   11775
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
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   9320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   2420
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   2300
      End
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   16935
      Begin VB.ComboBox txtTRPRS 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FormTransfer.frx":0000
         Left            =   7080
         List            =   "FormTransfer.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   240
         Width           =   3500
      End
      Begin VB.TextBox txtTrArea 
         BackColor       =   &H80000002&
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
         Left            =   15240
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtTrArea1 
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
         Left            =   10440
         TabIndex        =   31
         Top             =   240
         Width           =   6120
      End
      Begin VB.TextBox txtTrAreaRef 
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
         Left            =   7080
         TabIndex        =   30
         Top             =   240
         Width           =   3360
      End
      Begin VB.TextBox txtTrSupplierRef 
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
         Left            =   7080
         TabIndex        =   3
         Top             =   650
         Width           =   3360
      End
      Begin VB.TextBox txtTrPO 
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
         ForeColor       =   &H80000009&
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2750
      End
      Begin VB.ComboBox txtPOArea 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FormTransfer.frx":0034
         Left            =   2040
         List            =   "FormTransfer.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2750
      End
      Begin VB.TextBox txtTrNum 
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   13
         Top             =   240
         Width           =   2750
      End
      Begin VB.TextBox txtTrSupplier 
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
         Left            =   10440
         TabIndex        =   12
         Top             =   650
         Width           =   6105
      End
      Begin VB.TextBox txTrId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTrRemark 
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
         Left            =   7080
         TabIndex        =   4
         Top             =   1080
         Width           =   9500
      End
      Begin MSMask.MaskEdBox txtTrDate 
         Height          =   420
         Left            =   2040
         TabIndex        =   1
         Top             =   650
         Width           =   2750
         _ExtentX        =   4842
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
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
         BackColor       =   &H00C0E0FF&
         Caption         =   "AREA REF."
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
         Index           =   0
         Left            =   5520
         TabIndex        =   29
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "NUMBER"
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
         Left            =   360
         TabIndex        =   19
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "DATE"
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
         Left            =   360
         TabIndex        =   18
         Top             =   700
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "PO/CV NO."
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
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "SUPPLIER REF."
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
         Index           =   8
         Left            =   5520
         TabIndex        =   16
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Index           =   7
         Left            =   5520
         TabIndex        =   15
         Top             =   1095
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView lvwTr 
      Height          =   7260
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   12806
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
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
      Index           =   15
      Left            =   12000
      TabIndex        =   25
      Top             =   9840
      Width           =   1050
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   180
      Left            =   4200
      TabIndex        =   21
      Top             =   315
      Width           =   7650
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   20
      Top             =   90
      Width           =   7845
   End
End
Attribute VB_Name = "FormTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private POSearchLI                 As ListItem

Public TrRemark, TRArea As String
Private Items(7) As String
Private TrId, EncodeMode, ItemName, ItemCode, ItemGroup, TrResult, MRSInfo       As String
Private ItemCodeDel, ItemNoDel, ItemQtyDel, TxtVal, NumVal                       As String
Private TrItemNo, EndCost, EndAmount, EndQty, TrTotal, ItemAmount, POTotal       As Double
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwTr
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub CommandExecute()
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    'FormMainMenu.cmdExit.SetFocus
    lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdNew_Click()
ClearBox
ClearItemBox
ClearFrame
ButtonState False
EncodeMode = "A": cmdAdd.Enabled = True: cmdCancel.Enabled = True
frameOpt.Visible = True: frameOpt.Left = 7600: frameOpt.Top = 2650

lvwTr.ListItems.Clear
     TrId = Format$(GetNextPOID, "000000"): TrItemNo = 0: txtTrNum.Text = TrId: txtTrDate = Format$(Now, "mm/dd/yyyy")
     DeleteTemporary
            strsql = " Insert Into InventoryTemp Select * From Inventory": CommandExecute
End Sub
Private Sub cmdCV_Click()
    BoxState True
    lblOpt.Caption = "CV": frameOpt.Visible = False: txtTrDate.SetFocus: txtTrPO.Locked = False
    txtTrPO.SetFocus: txtTrPO.Text = "CV": txtTrPO.SelStart = 3
End Sub
Private Sub cmdPO_Click()
    BoxState True
    lblOpt.Caption = "PO": frameOpt.Visible = False: txtTrDate.SetFocus: txtTrPO.Locked = True
End Sub
Private Sub CmdAdd_Click()
    'EncodeMode = "A"
    'ClearFrame
    'ClearItemBox
    'LoadItemsList
End Sub
Private Sub cmdCancel_Click()
    DeleteTemporary
    ClearBox
    ClearItemBox
    ClearFrame
    BoxState False
    ButtonState False
    cmdNew.Enabled = True
    cmdSearch.Enabled = True
    lvwTr.ListItems.Clear
    cmdNew.SetFocus
End Sub
Private Sub cmdSearch_Click()
   ClearBox
   ClearItemBox
   ClearFrame
   DeleteTemporary
   BoxState False
   'LoadSearchList
   'EncodeMode = "S": txtTrSearch.Text = "": cboTrSearch.SetFocus:
   'txtPOArea.Visible = False: txtPOPrs.Visible = True
End Sub
Private Sub cmdPrint_Click()
TRArea = txtTrArea1.Text
strsql = "Update TRDetailsTemp SET TRSupplierRef = '" & txtTrSupplierRef.Text & "', TRRemark = '" & txtTrRemark.Text & "' "
CommandExecute
InsertTrTotal
    'If EncodeMode = "A" Or EncodeMode = "E" Then
    '  strsql = " Insert Into PODetails Select * From PODetailsTemp "
    '  CommandExecute
    'End If
    ClearBox
    ClearItemBox
    ClearFrame
    BoxState False
    ButtonState False
    cmdNew.Enabled = True
    cmdSearch.Enabled = True
    
   Load DataEnvironment1
   If DataEnvironment1.rsTransHeader.State <> 0 Then DataEnvironment1.rsTransHeader.Close
     ReportTransfer.Refresh
   If ReportTransfer.Visible = False Then ReportTransfer.Show
   
    If EncodeMode = "S" Then
        DeleteTemporary
    End If
Set mmsADORst = Nothing
End Sub
Private Sub cmdDelete_Click()
  strsql = "Delete From TRDetailsTemp WHERE TRItemNo like '" & txtTrItemNo.Text & "'"
  CommandExecute
  frameItemAdd.Visible = False: lvwTr.SetFocus
  LoadEditTemp
End Sub
Private Sub cmdSaveTRDetails_Click()
GetItemAmount
strsql = "Update TRDetailsTemp SET TRQty = '" & txtTrQty.Text & "', TRAmount = '" & txtTrAmount.Text & "' " _
         & "WHERE TRItemNo like '" & txtTrItemNo.Text & "'"
CommandExecute
frameItemAdd.Visible = False: lvwTr.SetFocus
LoadEditTemp
End Sub
Private Sub LoadEditTemp()
Dim TrLI  As ListItem
On Error GoTo LocalError
    strsql = "Select * from TRDetailsTemp": CommandExecute
    lvwTr.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set TrLI = lvwTr.ListItems.Add(, , !TrItemNo & "")
            TrLI.SubItems(1) = !TRCode & ""
            TrLI.SubItems(2) = !TRItem & ""
            TrLI.SubItems(3) = !TRQty & ""
            TrLI.SubItems(4) = !TRUnit & ""
            TrLI.SubItems(5) = Format$(!TRCost, "#,###.#0") & ""
            TrLI.SubItems(6) = Format$(!TRAmount, "#,###.#0") & ""
            .MoveNext
        Loop
      End With
      GetTrTotal
      InsertTrTotal
LocalError:
    Exit Sub
End Sub
Private Sub GetTrTotal()
    strsql = " SELECT SUM(TRAmount) as SubTotal FROM TRDetailsTemp "
    CommandExecute
    TrTotal = mmsADORst.Fields!Subtotal
    txtTrTotal.Text = Format$(TrTotal, "#,###.#0")
End Sub
Private Sub InsertTrTotal()
On Error GoTo LocalError
        strsql = "Update TRDetailsTemp SET TRTotal = '" & txtTrTotal.Text & "'"
        CommandExecute
LocalError:
    Exit Sub
End Sub
'---------------------------------------------------------------------------------
'                                   T E X T B O X   E V E N T S
'---------------------------------------------------------------------------------

Private Sub txtTrDate_GotFocus()
   txtTrDate.SelLength = Len(txtTrDate.Text)
End Sub
Private Sub txtTrDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtTrPO.SetFocus
   End If
End Sub
Private Sub txtTrPO_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      'txtTrArea.SetFocus
    Else
      txtTrPO_DblClick
    End If
End Sub
Private Sub txtTrPO_DblClick()
If lblOpt.Caption = "PO" Then
    'BoxState False
    SetlvwTRPO
    lvwTr.ListItems.Clear
    LoadPOList
End If
End Sub
Private Sub LoadPOList()
On Error GoTo LocalError
    strsql = "SELECT DISTINCT PONum, PODate ,POSupplier, POPrs, POArea, POTotal FROM PODetails ORDER BY PONum ": CommandExecute
    lvwPO.ListItems.Clear
    LoadPO
LocalError:
    Exit Sub
End Sub
Private Sub LoadPO()
    With mmsADORst
        Do Until .EOF
            Set POSearchLI = lvwPO.ListItems.Add(, , !PONum & "")
            POSearchLI.SubItems(1) = !PODate & ""
            POSearchLI.SubItems(2) = !POSupplier & ""
            POSearchLI.SubItems(3) = !POPrs & ""
            POSearchLI.SubItems(4) = !POArea & ""
            POSearchLI.SubItems(5) = !POTotal & ""
           .MoveNext
        Loop
     End With
End Sub
'------------------------------------------------------------------------------------
'                  P - O   S  E  A  R  C  H
'-----------------------------------------------------------------------------------
Private Sub txtPOSearch_Change()
   strsql = "Select DISTINCT PONum, PODate ,POSupplier, POPrs, POArea, POTotal from PODetails where POID like '" & txtPOSearch.Text & "%'" & "Order by PONum"
   CommandExecute
   lvwPO.ListItems.Clear
   LoadPO
End Sub
Private Sub txtPOSearch_Click()
    txtPOSearch.Text = ""
End Sub
Private Sub txtPOSearch_GotFocus()
   txtPOSearch.Text = ""
   txtPOSearch.SelLength = Len(txtPOSearch.Text)
End Sub
Private Sub txtPOSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwPO.SetFocus
   End If
End Sub
'------------   P O    L I S T V I E W   -------------------
Private Sub lvwPO_KeyPress(KeyAscii As Integer)
    lvwPO_DblClick
End Sub
Private Sub lvwPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtTrPO.Text = lvwPO.SelectedItem.Text
      txtTrSupplier.Text = lvwPO.SelectedItem.SubItems(2)
      txtTrAreaRef.Text = lvwPO.SelectedItem.SubItems(3)
      txtTrArea.Text = lvwPO.SelectedItem.SubItems(4)
      txtTrArea1.Text = GetArea(txtTrArea.Text)
      txtTrTotal.Text = lvwPO.SelectedItem.SubItems(5)
    End With
    'LoadPODetails
End Sub
Private Function GetArea(ByVal area As String) As String
    Select Case area
        Case "DCO"
           GetArea = "139 PEACOCK ST ECOLAND SUBD. DAVAO CITY"
        Case "B1"
           GetArea = "BARACATAN, STA. CRUZ DAVAO DEL SUR"
        Case "B2"
           GetArea = "BATOBATO, SAN ISIDRO DAVAO ORIENTAL"
        Case "B3"
           GetArea = "MALUNGON, SARANGANI PROVINCE"
        Case "B4"
           GetArea = "STA TERESITA, BAYUGAN AGUSAN DEL SUR"
    End Select
End Function
Private Sub lvwPO_DblClick()
    BoxState True
    framePO.Visible = False: cmdPrint.Enabled = True: lvwTr.Enabled = True: txtTrSupplierRef.SetFocus
    LoadPODetails
    InsertTrTotal
End Sub
Private Sub LoadPODetails()
Dim TrLI  As ListItem
On Error GoTo LocalError
    strsql = "Select * from PODetails where PONum like '" & txtTrPO.Text & "' Order by POItemNo"
    CommandExecute
    lvwPO.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set TrLI = lvwTr.ListItems.Add(, , !POItemNo & ""): Items(0) = !POItemNo
            TrLI.SubItems(1) = !POCode & "": Items(1) = !POCode
            TrLI.SubItems(2) = !POItem & "": Items(2) = !POItem
            TrLI.SubItems(3) = !POQty & "": Items(3) = !POQty
            TrLI.SubItems(4) = !POUnit & "": Items(4) = !POUnit
            TrLI.SubItems(5) = Format$(!POCost, "#,###.#0") & "": Items(5) = Format$(!POCost, "#,###.#0")
            TrLI.SubItems(6) = Format$(!POAmount, "#,###.#0") & "": Items(6) = Format$(!POAmount, "#,###.#0")
            Items(7) = !POGroup
            LoadTempTR
            .MoveNext
        Loop
      End With
      'GetPOTotal
      'txtPOTotal.Text = Format$(POTotal, "#,###.#0")
      'InsertPOTotal
LocalError:
    Exit Sub
End Sub
Private Sub LoadTempTR()
On Error GoTo LocalError

  If EncodeMode = "A" Then
         strsql = "INSERT INTO TRDetailsTemp (  TRID"
         strsql = strsql & "            , TRItemNo"
         strsql = strsql & "            , TRNum"
         strsql = strsql & "            , TRDate"
         strsql = strsql & "            , TRPo"
         strsql = strsql & "            , TRArea"
         strsql = strsql & "            , TRAreaRef"
         strsql = strsql & "            , TRSupplier"
         strsql = strsql & "            , TRSupplierRef"
         strsql = strsql & "            , TRRemark"
         strsql = strsql & "            , TRGroup"
         strsql = strsql & "            , TRCode"
         strsql = strsql & "            , TRItem"
         strsql = strsql & "            , TRQty"
         strsql = strsql & "            , TRUnit"
         strsql = strsql & "            , TRCost"
         strsql = strsql & "            , TRAmount"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & TrId
         strsql = strsql & ", '" & Items(0) & "'"
         strsql = strsql & ", '" & Replace$(txtTrNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrPO.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrArea.Text, "'", "''") & "'"
          strsql = strsql & ", '" & Replace$(txtTrAreaRef.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrSupplier.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrSupplierRef.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtTrRemark.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Items(7) & "'"   'GROUP
         strsql = strsql & ", '" & Items(1) & "'"
         strsql = strsql & ", '" & Items(2) & "'"
         strsql = strsql & ", '" & Items(3) & "'"
         strsql = strsql & ", '" & Items(4) & "'"
         strsql = strsql & ", '" & Items(5) & "'"
         strsql = strsql & ", '" & Items(6) & "'"
         strsql = strsql & ")"
        mmsAdoCmd.CommandText = strsql: Set mmsADORst = mmsAdoCmd.Execute
  End If
  
LocalError:
    Exit Sub
End Sub

'-----------------------------------------------------------------------------
Private Sub txtTrSupplierRef_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtTrRemark.SetFocus
   End If
End Sub
Private Sub txtTrRemark_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      lvwTr.SetFocus
   End If
End Sub
Private Sub txtTrQty_GotFocus()
 txtTrQty.SelLength = Len(txtTrQty.Text)
End Sub
Private Sub txtTrQty_Change()
  TxtVal = txtTrQty.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtTrQty.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtTrQty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       cmdSaveTRDetails.SetFocus
       GetItemAmount
   End If
End Sub
Private Sub GetItemAmount()
    txtTrQty.Text = Format$(txtTrQty, "#,###.#0")
    ItemAmount = CDbl(txtTrCost.Text) * CDbl(txtTrQty.Text)
    txtTrAmount.Text = Format$(ItemAmount, "#,###.#0")
    txtTrCost.Text = Format$(txtTrCost, "#,###.#0")
End Sub
'----------------------------------------
                     
'-----------------------------------------


'------------------------------------------------------------------------------------
'                      S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'                  I T E M S   C O N T R O L S
'-----------------------------------------------------------------------------------


'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub SetlvwTr()
lvwTr.ColumnHeaders.Clear
    With lvwTr
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " # ", .Width * 0.03
        .ColumnHeaders.Add , , "CODE", .Width * 0.07
        .ColumnHeaders.Add , , "ITEM", .Width * 0.35
        .ColumnHeaders.Add , , "QTY", Width * 0.08
        .ColumnHeaders.Add , , "UNIT", .Width * 0.05
        .ColumnHeaders.Add , , "U/P", Width * 0.12
        .ColumnHeaders.Add , , "AMOUNT", Width * 0.12
    End With
  lvwTr.ColumnHeaders.Item(4).Alignment = lvwColumnRight:
  lvwTr.ColumnHeaders.Item(6).Alignment = lvwColumnRight
  lvwTr.ColumnHeaders.Item(7).Alignment = lvwColumnRight
End Sub
Private Sub lvwTR_DblClick()
   If lvwTr.SelectedItem Is Nothing Then
   Else
      frameItemAdd.Visible = True: txtTrQty.SetFocus
      txtTrItemNo.Text = lvwTr.SelectedItem: txtTrItem.Text = lvwTr.SelectedItem.SubItems(2)
      txtTrUnit.Text = lvwTr.SelectedItem.SubItems(4): txtTrCost.Text = lvwTr.SelectedItem.SubItems(5): txtTrAmount.Text = lvwTr.SelectedItem.SubItems(6)
      txtTrQty.Text = lvwTr.SelectedItem.SubItems(3)
   End If
End Sub
'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "select * from MRSDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub LoadTrDetails()
Dim TrLI  As ListItem
On Error GoTo LocalError
      'GetPOTotal
      '  txtPOTotal.Text = Format$(POTotal, "#,###.#0")
      'InsertPOTotal
    strsql = "SELECT * from PODetailsTemp ORDER BY POItemNo "
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwTr.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set TrLI = lvwTr.ListItems.Add(, , !POItemNo & "")
            TrLI.SubItems(1) = !POCode & ""
            TrLI.SubItems(2) = !POItem & ""
            TrLI.SubItems(3) = !POQty & ""
            TrLI.SubItems(4) = !POUnit & ""
            TrLI.SubItems(5) = Format$(!POCost, "#,###.#0") & ""
            TrLI.SubItems(6) = Format$(!POAmount, "#,###.#0") & ""
            .MoveNext
        Loop
      End With

    cmdAdd.SetFocus
LocalError:
    Exit Sub
End Sub
Private Function GetNextPOID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(POID) AS MaxID FROM PODetails"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextPOID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextPOID = 1
       Else
           GetNextPOID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtTrDate.Text = "" Then
        MsgBox "Fill-up PO Date.", vbExclamation, "PO Date Required"
        txtTrDate.SetFocus: Exit Function
    End If
    DataValidation = True
End Function
Private Function DataItemValidation() As Boolean
On Error GoTo LocalError
  DataItemValidation = False
    'If txtPOItem.Text = "" Or txtPOUnit.Text = "" Then
    '    MsgBox "No Item Name.", vbExclamation, "Item Required"
    '    txtPOItem.SetFocus: Exit Function
    'End If
    DataItemValidation = True
LocalError:
    'MsgBox "Check Data", vbExclamation, "Check"
    Exit Function
End Function
Private Sub DeleteTemporary()
On Error GoTo LocalError
    strsql = "Delete From TRDetailsTemp": CommandExecute
    Set mmsADORst = Nothing
LocalError:
    Exit Sub
End Sub
Private Sub ClearBox()
    lvwTr.ListItems.Clear
    txtTrNum.Text = "": txtTrDate.Text = "__/__/____": txtTrPO.Text = ""
    txtTrAreaRef.Text = "": txtTrArea1.Text = "": txtTrSupplier.Text = "": txtTrSupplierRef.Text = "": txtTrRemark.Text = ""
    txtTrTotal.Text = ""
End Sub
Private Sub ClearItemBox()
    'ItemCode = "": ItemGroup = "": txtTrTotal.Text = ""
    'txtPOItem.Text = "": txtPOGroup.Text = "": txtPOUnit.Text = ""
    'txtPOQty.Text = "": txtPOCost.Text = "": txtPOAmount.Text = ""
End Sub
Private Sub ClearFrame()
    framePO.Visible = False: frameItemAdd.Visible = False: frameOpt.Visible = False
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtTrDate.Enabled = boxEnabled
    txtTrPO.Enabled = boxEnabled
    txtTrSupplierRef.Enabled = boxEnabled
    txtTrRemark.Enabled = boxEnabled
End Sub
Private Sub ItemBoxState(boxEnabled As Boolean)
    'txtTrItem.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwTr.Enabled = buttonEnabled
    cmdNew.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdSearch.Enabled = buttonEnabled
    cmdPrint.Enabled = buttonEnabled
    'cmdEdit.Enabled = buttonEnabled
    'cmdDelete.Enabled = buttonEnabled
End Sub


'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HC0E0FF
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0E0FF
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HC0E0FF
End Sub
Private Sub cmdSearch_LostFocus()
   cmdSearch.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0E0FF
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdPrint_GotFocus()
   cmdPrint.BackColor = &HC0E0FF
End Sub
Private Sub cmdPrint_LostFocus()
   cmdPrint.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0E0FF
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub cmdSaveTRDetails_GotFocus()
   cmdSaveTRDetails.BackColor = &HC0E0FF
End Sub
Private Sub cmdSaveTRDetails_LostFocus()
   cmdSaveTRDetails.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0E0FF
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub


