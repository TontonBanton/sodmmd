VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormPRS 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  PURCHASE REQUISITION"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   Icon            =   "PRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameItemAdd 
      BackColor       =   &H00C0FFFF&
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
      Height          =   7695
      Left            =   840
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   7045
      Begin VB.ComboBox txtPRSTransact 
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
         ItemData        =   "PRS.frx":1E72
         Left            =   1800
         List            =   "PRS.frx":1E74
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3130
         Width           =   3255
      End
      Begin VB.ComboBox txtPRSSeason 
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
         ItemData        =   "PRS.frx":1E76
         Left            =   5040
         List            =   "PRS.frx":1E83
         TabIndex        =   27
         Text            =   "-"
         Top             =   3130
         Width           =   1660
      End
      Begin VB.ComboBox txtPRSDept 
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
         ItemData        =   "PRS.frx":1E9D
         Left            =   1800
         List            =   "PRS.frx":1E9F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2600
         Width           =   4905
      End
      Begin VB.ComboBox txtPRSDept1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         ItemData        =   "PRS.frx":1EA1
         Left            =   360
         List            =   "PRS.frx":1EA3
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   6600
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtPRSInv 
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
         Height          =   650
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   31
         Top             =   5760
         Width           =   4900
      End
      Begin VB.TextBox txtPRSServed 
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
         Height          =   650
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   30
         Top             =   5160
         Width           =   4900
      End
      Begin VB.TextBox txtPRSBudget 
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
         Height          =   650
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   29
         Top             =   4560
         Width           =   4900
      End
      Begin VB.CommandButton cmdSavePRSDetails 
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
         Height          =   630
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6720
         Width           =   4900
      End
      Begin VB.TextBox txtPRSItemCode 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   300
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1050
         Width           =   6405
      End
      Begin VB.TextBox txtPRSUnit 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1950
         Width           =   4900
      End
      Begin VB.TextBox txtPRSItem 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   300
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   450
         Width           =   6400
      End
      Begin VB.TextBox txtPRSQty 
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
         Height          =   650
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3960
         Width           =   4900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "DETAILS:"
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
         Index           =   11
         Left            =   300
         TabIndex        =   48
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "UNIT"
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
         Index           =   10
         Left            =   300
         TabIndex        =   47
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "INVENTORY"
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
         Index           =   9
         Left            =   360
         TabIndex        =   46
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "SERVED"
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
         Left            =   360
         TabIndex        =   45
         Top             =   5280
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "BUDGETED"
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
         Left            =   360
         TabIndex        =   44
         Top             =   4680
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "REQUEST"
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
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   4080
         Width           =   1110
      End
   End
   Begin VB.Frame frameItems 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1755
      Left            =   8760
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   3465
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
         Left            =   13200
         TabIndex        =   20
         Top             =   250
         Width           =   3500
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "ITEM LIBRARY"
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
         Left            =   9700
         TabIndex        =   19
         Top             =   250
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   18
         Top             =   280
         Width           =   9500
      End
      Begin MSComctlLib.ListView lvwPRSItems 
         Height          =   6045
         Left            =   0
         TabIndex        =   21
         Top             =   960
         Width           =   16900
         _ExtentX        =   29819
         _ExtentY        =   10663
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
      Height          =   461
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   1550
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
      Height          =   461
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   1550
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   858
      Left            =   0
      TabIndex        =   37
      Top             =   9840
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
         Left            =   4730
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   9320
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
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   2300
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&XIT"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   2800
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
      Height          =   1170
      Left            =   8760
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtPRSSearch 
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   35
         Top             =   240
         Width           =   5000
      End
      Begin VB.ComboBox cboPRSSearch 
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
         ItemData        =   "PRS.frx":1EA5
         Left            =   120
         List            =   "PRS.frx":1EB5
         TabIndex        =   34
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.ListView lvwPRSSearch 
         Height          =   6000
         Left            =   0
         TabIndex        =   36
         Top             =   840
         Width           =   16900
         _ExtentX        =   29819
         _ExtentY        =   10583
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1212
      Left            =   0
      TabIndex        =   7
      Top             =   770
      Width           =   16935
      Begin VB.TextBox txtPRSFarm 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   15960
         TabIndex        =   49
         Text            =   "B3-MALUNGON"
         Top             =   120
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtPRSRemark 
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
         Left            =   6840
         TabIndex        =   11
         Top             =   600
         Width           =   9800
      End
      Begin VB.TextBox txtPRSNum 
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
         TabIndex        =   8
         Top             =   200
         Width           =   2500
      End
      Begin MSMask.MaskEdBox txtPRSDate 
         Height          =   450
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   794
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
      Begin MSMask.MaskEdBox txtPRSNeed 
         Height          =   450
         Left            =   6840
         TabIndex        =   10
         Top             =   200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   794
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "REMARKS "
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
         Left            =   5160
         TabIndex        =   40
         Top             =   650
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "DATE NEEDED"
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
         Index           =   1
         Left            =   5160
         TabIndex        =   14
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "PRS DATE"
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
         Left            =   240
         TabIndex        =   13
         Top             =   650
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "PRS NUMBER"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView lvwPRS 
      Height          =   7740
      Left            =   0
      TabIndex        =   15
      Top             =   1920
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   13653
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
      Left            =   6840
      TabIndex        =   39
      Top             =   120
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
      Left            =   6840
      TabIndex        =   38
      Top             =   360
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   5940
      Picture         =   "PRS.frx":1ED6
      Stretch         =   -1  'True
      Top             =   45
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "FormPRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String

Private PRSId, EncodeMode, ItemCode, ItemGroup, PRSResult, TxtVal, NumVal, ItemNoDel, Details  As String
Private ItemStock, ItemAmount, PRSItemNo                                                       As Double
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwPRS
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub CommandExecute()
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdNew_Click()
Dim PRSGroup As String
    EncodeMode = "A"
    ButtonState False
    cmdAdd.Enabled = True
    cmdCancel.Enabled = True
    BoxState True
    ClearBox
    ClearItemBox
    ClearFrame
    lvwPRS.ListItems.Clear
     PRSGroup = FormMainMenu.lblArea.Caption & "-" & "PRS"
     PRSId = Format$(GetNextPRSID, "000000")
     PRSItemNo = 0
     txtPRSDate = Format$(Now, "mm/dd/yyyy")
     txtPRSNeed = Format$(Now, "mm/dd/yyyy")
     txtPRSNum.Text = PRSGroup & "-" & PRSId
     txtPRSDate.SetFocus
     DeleteTemporary
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ClearFrame
    ClearItemBox
    LoadItemsList
End Sub
Private Sub cmdSavePRSDetails_Click()
   If Not DataItemValidation Then
        Exit Sub
   End If
   txtPRSFarm.Text = "B3-Malungon"
   PRSItemNo = PRSItemNo + 1
   FormatText
   SaveTemporary
   LoadPRSDetails
   ClearFrame
   cmdPrint.Enabled = True
   lvwPRS.Enabled = True
End Sub
Private Sub cmdEdit_Click()
   'MsgBox "Edit"
End Sub
Private Sub cmdDelete_Click()
    'mmsAdoCmd.CommandText = "Delete From PRSDetails"
    'Set mmsADORst = mmsAdoCmd.Execute
    'MsgBox "delete database"
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
    lvwPRS.ListItems.Clear
    cmdNew.SetFocus
End Sub
Private Sub cmdSearch_Click()
   EncodeMode = "S"
   ClearBox
   ClearItemBox
   ClearFrame
   DeleteTemporary
   BoxState False
   LoadSearchList
   txtPRSSearch.Text = ""
   cboPRSSearch.SetFocus
End Sub
Private Sub cmdPrint_Click()
    If EncodeMode = "A" Or EncodeMode = "E" Then
      strsql = " Insert Into PRSDetails Select * From PRSDetailsTemp "
      CommandExecute
    End If
    
    ClearBox
    ClearItemBox
    ClearFrame
    BoxState False
    ButtonState False
    cmdNew.Enabled = True
    cmdSearch.Enabled = True
   
   Load DataEnvironment1
   If DataEnvironment1.rsCommand5.State <> 0 Then DataEnvironment1.rsCommand5.Close
     ReportPRS.Refresh
   If ReportPRS.Visible = False Then ReportPRS.Show
   
      If EncodeMode = "S" Then
        DeleteTemporary
    End If
    
Set mmsADORst = Nothing
End Sub
Private Sub cmdAddClose_Click()
    frameItems.Visible = False
    cmdAdd.SetFocus
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     ClearBox
     ClearItemBox
     ClearFrame
     Unload Me
     FormMainMenu.Show
     FormMainMenu.cmdExit.SetFocus
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
Private Sub cmdAddCCC_Click()
       'FormCCC.Show
End Sub
Private Sub cmdAddItem_Click()
       FormItems.Show
End Sub

Private Sub txtMISTransact_Change()

End Sub

'---------------------------------------------------------------------------------
'                              T E X T B O X    E V E N T S
'---------------------------------------------------------------------------------
Private Sub TxtPRSdate_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      SendKeys_
    ElseIf KeyAscii = 13 Then
      txtPRSNeed.SetFocus
    End If
End Sub
Private Sub txtPRSDate_GotFocus()
   txtPRSDate.SelLength = Len(txtPRSDate.Text)
End Sub
Private Sub TxtPRSNeed_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      SendKeys_
    ElseIf KeyAscii = 13 Then
      txtPRSRemark.SetFocus
    End If
End Sub
Private Sub txtPRSNeed_GotFocus()
  If Not IsDate(txtPRSDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtPRSDate.SetFocus
        txtPRSDate = Format$(Now, "mm/dd/yyyy")
  End If
   txtPRSNeed.SelLength = Len(txtPRSNeed.Text)
End Sub
Private Sub txtPRSRemark_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      LoadItemsList
    End If
    If Not IsDate(txtPRSDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtPRSDate.SetFocus
        txtPRSDate = Format$(Now, "mm/dd/yyyy")
    End If
End Sub
Private Sub txtPRSRemark_GotFocus()
   txtPRSRemark.SelLength = Len(txtPRSRemark.Text)
End Sub

'----------------- D E T A I L S
Private Sub txtPRSUnit_GotFocus()
   txtPRSUnit.SelLength = Len(txtPRSUnit.Text)
End Sub
Private Sub txtPRSUnit_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       txtPRSDept.SetFocus
    End If
End Sub
Private Sub txtPRSDept_GotFocus()
  strsql = "SELECT * FROM Area ORDER BY AreaID"
    CommandExecute
    With mmsADORst
    txtPRSDept.Clear
    Do While Not .EOF
        txtPRSDept.AddItem ![AreaName]: .MoveNext
    Loop
    txtPRSDept.ListIndex = 0
    End With
End Sub
Private Sub txtPRSDept_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       txtPRSTransact.SetFocus
    End If
End Sub
Private Sub txtPRSTransact_GotFocus()
    strsql = "SELECT * FROM Transact ORDER BY TransactID"
    CommandExecute
    With mmsADORst
    txtPRSTransact.Clear
    Do While Not .EOF
        txtPRSTransact.AddItem ![TransactName]: .MoveNext
    Loop
    txtPRSTransact.ListIndex = 0
    End With
End Sub
Private Sub txtprsTransact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtPRSSeason.SetFocus
End If
End Sub
Private Sub txtPRSSeason_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtPRSQty.SetFocus
End If
End Sub
Private Sub txtPRSQty_GotFocus()
  txtPRSQty.SelLength = Len(txtPRSQty.Text)
End Sub
Private Sub txtPRSQty_Change()
  TxtVal = txtPRSQty.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPRSQty.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtPRSQty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtPRSBudget = txtPRSQty.Text
       txtPRSBudget.SetFocus
   End If
End Sub
Private Sub txtPRSBudget_GotFocus()
   txtPRSBudget.SelLength = Len(txtPRSBudget.Text)
End Sub
Private Sub txtPRSBudget_Change()
  TxtVal = txtPRSBudget.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPRSBudget.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtPRSBudget_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtPRSServed = "0"
       txtPRSServed.SetFocus
   End If
End Sub
Private Sub txtPRSServed_GotFocus()
   txtPRSServed.SelLength = Len(txtPRSServed.Text)
End Sub
Private Sub txtPRSServed_Change()
  TxtVal = txtPRSServed.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPRSServed.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtPRSServed_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtPRSInv.SetFocus
   End If
End Sub
Private Sub txtPRSInv_GotFocus()
On Error GoTo LocalError
    strsql = "Select * From Inventory where ItemCode like '" & ItemCode & "' "
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
       
       txtPRSInv.Text = Format$(mmsADORst.Fields("AvailStock"))
       txtPRSInv.SelLength = Len(txtPRSInv.Text)
LocalError:
    txtPRSInv.Text = "0"
    txtPRSInv.SelLength = Len(txtPRSInv.Text)
    Exit Sub
End Sub
Private Sub txtPRSInv_Change()
  TxtVal = txtPRSInv.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPRSInv.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtPRSInv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       cmdSavePRSDetails.SetFocus
   End If
End Sub
'-----------------------------------------------------------------------------------
'                      S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboPRSSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtPRSSearch.SetFocus
   End If
   If KeyAscii < 255 Then
      SendKeys_
   End If
End Sub
Private Sub cboPRSSearch_GotFocus()
  cboPRSSearch.Text = "NUMBER"
End Sub
Private Sub txtPRSSearch_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     lvwPRSSearch.SetFocus
  End If
End Sub
Private Sub txtPRSSearch_Change()
Dim PRSSearchLI       As ListItem
On Error GoTo LocalError
    If cboPRSSearch.Text = "NUMBER" Then
       strsql = "Select * from PRSDetails where PRSID like '" & txtPRSSearch.Text & "%'" & "Order by PRSNum"
    ElseIf cboPRSSearch.Text = "DATE" Then
       strsql = "Select * from PRSDetails where PRSDate like '" & txtPRSSearch.Text & "%'" & "Order by PRSNum"
    ElseIf cboPRSSearch.Text = "ITEM" Then
       strsql = "Select * from PRSDetails where PRSItem like '" & txtPRSSearch.Text & "%'" & "Order by PRSNum"
    ElseIf cboPRSSearch.Text = "REMARKS" Then
       strsql = "Select * from PRSDetails where PRSRemark like '%" & txtPRSSearch.Text & "%'" & "Order by PRSNum"
    End If
    
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LoadPRSSearch

LocalError:
    Exit Sub
End Sub

'-----------------------  S E A R C H  L I S T V I E W ----------------------
Private Sub lvwPRSSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      SearchLoad
    End With
End Sub
Private Sub lvwPRSSearch_DblClick()
   frameSearch.Visible = False
   BoxState False
   SaveTemporary
   LoadPRSDetails
   cmdPrint.Enabled = True
End Sub
Private Sub lvwPRSSearch_KeyPress(KeyAscii As Integer)
On Error GoTo LocalError
   If KeyAscii = 13 Then
       lvwPRSSearch_DblClick
       PRSResult = lvwPRSSearch.SelectedItem.Text
   End If
LocalError:
    Exit Sub
End Sub
Private Sub lvwPRSSearch_GotFocus()
On Error GoTo LocalError
      SearchLoad
LocalError:
    Exit Sub
End Sub
Private Sub SearchLoad()
      txtPRSNum.Text = lvwPRSSearch.SelectedItem.Text
      txtPRSDate.Text = Format$(lvwPRSSearch.SelectedItem.SubItems(1), "mm/dd/yyyy")
      txtPRSNeed.Text = Format$(lvwPRSSearch.SelectedItem.SubItems(2), "mm/dd/yyyy")
      'txtPRSDept.Text = lvwPRSSearch.SelectedItem.SubItems(4)
      txtPRSRemark.Text = lvwPRSSearch.SelectedItem.SubItems(5)
End Sub
Private Sub LoadSearchList()

   frameSearch.Top = lvwPRS.Top: frameSearch.Left = lvwPRS.Left: frameSearch.Height = lvwPRS.Height: frameSearch.Width = lvwPRS.Width
   frameSearch.Visible = True: lvwPRSSearch.ColumnHeaders.Clear
    With lvwPRSSearch
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NUMBER", .Width * 0.1
        .ColumnHeaders.Add , , "DATE", .Width * 0.08
        .ColumnHeaders.Add , , "NEEDED", .Width * 0.08
        .ColumnHeaders.Add , , "ITEM", .Width * 0.3
        .ColumnHeaders.Add , , "AREA", .Width * 0.18
        .ColumnHeaders.Add , , "REMARKS", .Width * 0.25
    End With
    lvwPRSSearch.Height = frameSearch.Height - 1000

    strsql = "SELECT DISTINCT PRSNum, PRSDate, PRSNeed, PRSItem" _
           & "     , PRSDept, PRSRemark" _
           & "     FROM PRSDetails ORDER BY PRSNum "
    CommandExecute
    
    LoadPRSSearch
    
End Sub
Private Sub LoadPRSSearch()
Dim PRSSearchLI   As ListItem
On Error GoTo LocalError

    lvwPRSSearch.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set PRSSearchLI = lvwPRSSearch.ListItems.Add(, , !PRSNum & "")
            PRSSearchLI.SubItems(1) = !PRSDate & ""
            PRSSearchLI.SubItems(2) = !PRSNeed & ""
            PRSSearchLI.SubItems(3) = !PRSItem & ""
            PRSSearchLI.SubItems(4) = !PRSDept & ""
            PRSSearchLI.SubItems(5) = !PRSRemark & ""
            .MoveNext
        Loop
    End With
     
    Set PRSSearchLI = Nothing
    Set mmsADORst = Nothing
    
LocalError:
    Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                      I T E M   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtItemSearch_Change()
Dim MaterialsLI  As ListItem
    strsql = "Select * from Materials where ItemName like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemName"
             
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwPRSItems.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwPRSItems.ListItems.Add(, , !ItemGroup & "")
            MaterialsLI.SubItems(1) = !ItemCode & ""
            MaterialsLI.SubItems(2) = !ItemName & ""
            MaterialsLI.SubItems(4) = !Unit & ""
            MaterialsLI.SubItems(5) = CStr(!ItemId)
            .MoveNext
        Loop
     End With
    Set MaterialsLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      lvwPRSItems.SetFocus
    End If
End Sub
Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   'If KeyCode = 112 Then
   '    cmdAddItem_Click
   'End If
End Sub
Private Sub txtItemSearch_Click()
    txtItemSearch.Text = ""
End Sub
Private Sub txtItemSearch_GotFocus()
   txtItemSearch.Text = ""
   txtItemSearch.SelLength = Len(txtItemSearch.Text)
End Sub
'---------- I T E M S   L I S T V I E W ----------
Private Sub lvwPRSItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      ItemGroup = lvwPRSItems.SelectedItem
      ItemCode = lvwPRSItems.SelectedItem.SubItems(1)
      LoadItemSearch
    End With
End Sub
Private Sub lvwPRSItems_DblClick()
    frameItems.Visible = False
    frameItemAdd.Top = lvwPRS.Top - 100: frameItemAdd.Left = 7500: frameItemAdd.Height = lvwPRS.Height + 100: frameItemAdd.Width = 10500
    frameItemAdd.Visible = True
    BoxState False
    ItemBoxState True
    txtPRSDept.SetFocus
End Sub
Private Sub lvwPRSItems_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
      cmdAddItem_Click
   End If
End Sub
Private Sub lvwPRSItems_KeyPress(KeyAscii As Integer)
   lvwPRSItems_DblClick
End Sub
Private Sub lvwPRSItems_GotFocus()
On Error GoTo LocalError
      ItemGroup = lvwPRSItems.SelectedItem.Text
      ItemCode = lvwPRSItems.SelectedItem.SubItems(1)
      LoadItemSearch
LocalError:
    Exit Sub
End Sub
Private Sub LoadItemSearch()
      txtPRSItemCode.Text = ItemGroup & "-" & ItemCode
      txtPRSItem.Text = lvwPRSItems.SelectedItem.SubItems(2)
      txtPRSUnit.Text = lvwPRSItems.SelectedItem.SubItems(4)
End Sub
Private Sub LoadItemsList()
Dim MaterialsLI  As ListItem
On Error GoTo LocalError
    
    If Not DataValidation Then
       Exit Sub
    End If
    
       frameItems.Top = Frame1.Top: frameItems.Left = lvwPRS.Left: frameItems.Width = lvwPRS.Width: frameItems.Height = 8700
       lvwPRSItems.Top = 950: lvwPRSItems.Height = frameItems.Height - 800: frameItems.Visible = True: txtItemSearch.SetFocus

     lvwPRSItems.ColumnHeaders.Clear
     With lvwPRSItems
        .ColumnHeaders.Add , , "", .Width * 0.1
        .ColumnHeaders.Add , , "", .Width * 0.1
        .ColumnHeaders.Add , , "", .Width * 0.5
        .ColumnHeaders.Add , , "", .Width * 0#
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", Width * 0#
     End With
     lvwPRSItems.ColumnHeaders.Item(4).Alignment = lvwColumnRight
                             
    strsql = "SELECT * From Materials Order By ItemGroup, ItemName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwPRSItems.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwPRSItems.ListItems.Add(, , !ItemGroup & "")
            MaterialsLI.SubItems(1) = !ItemCode & ""
            MaterialsLI.SubItems(2) = !ItemName & ""
            MaterialsLI.SubItems(4) = !Unit & ""
            MaterialsLI.SubItems(5) = CStr(!ItemId)
            .MoveNext
        Loop
     End With
    Set MaterialsLI = Nothing
    Set mmsADORst = Nothing
    

LocalError:
    Exit Sub
End Sub
'-------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub lvwPRS_DblClick()
   If lvwPRS.SelectedItem Is Nothing Then
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
    ItemNoDel = lvwPRS.SelectedItem
    
        strsql = "Delete From PRSDetailsTemp where PRSItemNo like " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
    
        strsql = "Update PRSDetailsTemp SET PRSItemNo = PRSItemNo - 1"
        strsql = strsql & " where PRSItemNo > " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
        
        PRSItemNo = PRSItemNo - 1
    
    LoadPRSDetails
    
End Sub
'     ----------    S E T  L I S T V I E W   -------------
Private Sub SetlvwPRS()
lvwPRS.ColumnHeaders.Clear
    With lvwPRS
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " # ", .Width * 0.02
        .ColumnHeaders.Add , , "Area", .Width * 0.1
        .ColumnHeaders.Add , , " ", .Width * 0.2
        .ColumnHeaders.Add , , "Description", .Width * 0.35
        .ColumnHeaders.Add , , "Unit", .Width * 0.05
        .ColumnHeaders.Add , , "Budget", Width * 0.1
        .ColumnHeaders.Add , , "Served", Width * 0.1
        .ColumnHeaders.Add , , "Inventory", Width * 0.1
        .ColumnHeaders.Add , , "QTY", Width * 0.1
    End With
  lvwPRS.ColumnHeaders.Item(6).Alignment = lvwColumnRight
  lvwPRS.ColumnHeaders.Item(7).Alignment = lvwColumnRight
  lvwPRS.ColumnHeaders.Item(8).Alignment = lvwColumnRight
  lvwPRS.ColumnHeaders.Item(9).Alignment = lvwColumnRight
End Sub

'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------

Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "select * from PRSDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub LoadPRSDetails()
Dim PRSLI             As ListItem
On Error GoTo LocalError
 
    strsql = "SELECT * from PRSDetailsTemp ORDER BY PRSItemNo "
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwPRS.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set PRSLI = lvwPRS.ListItems.Add(, , !PRSItemNo & "")
            PRSLI.SubItems(1) = !PRSDept & ""
            PRSLI.SubItems(2) = !PRSTransact & " - " & !PRSSeason & ""
            PRSLI.SubItems(3) = !PRSItem & ""
            PRSLI.SubItems(4) = !PRSUnit & ""
            PRSLI.SubItems(5) = !PRSBudget & ""
            PRSLI.SubItems(6) = !PRSServed & ""
            PRSLI.SubItems(7) = !PRSInv & ""
            PRSLI.SubItems(8) = !PRSQty & ""
            .MoveNext
        Loop
      End With
    cmdAdd.SetFocus
LocalError:
    Exit Sub
End Sub
Private Sub SaveTemporary()
Dim strsql          As String
On Error GoTo LocalError
Details = txtPRSDept.Text & " - " & txtPRSTransact.Text & " - " & txtPRSSeason.Text
   If EncodeMode = "A" Then
         strsql = "INSERT INTO PRSDetailsTemp    (  PRSId"
         strsql = strsql & "            , PRSItemNo"
         strsql = strsql & "            , PRSNum"
         strsql = strsql & "            , PRSDate"
         strsql = strsql & "            , PRSNeed"
         strsql = strsql & "            , PRSFarm"
         strsql = strsql & "            , PRSRemark"
         strsql = strsql & "            , PRSDept"
         strsql = strsql & "            , PRSTransact"
         strsql = strsql & "            , PRSSeason "
         strsql = strsql & "            , PRSCode"
         strsql = strsql & "            , PRSItem"
         strsql = strsql & "            , PRSUnit"
         strsql = strsql & "            , PRSQty"
         strsql = strsql & "            , PRSBudget"
         strsql = strsql & "            , PRSServed"
         strsql = strsql & "            , PRSInv"
         strsql = strsql & "            , PRSDetails"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & PRSId
         strsql = strsql & ", '" & PRSItemNo & "'"
         strsql = strsql & ", '" & Replace$(txtPRSNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSNeed.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSFarm.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSRemark.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSDept.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSTransact.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSSeason.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtPRSItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSQty.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSBudget.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSServed.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPRSInv.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Details & "'"
         strsql = strsql & ")"

         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
   End If
   
   If EncodeMode = "S" Then
     strsql = " Insert Into PRSDetailsTemp Select * From PRSDetails Where PRSNum like '" & txtPRSNum.Text & "' Order By PRSItemNo "
     mmsAdoCmd.CommandText = strsql
     Set mmsADORst = mmsAdoCmd.Execute
     cmdCancel.Enabled = True
   End If
     
    Set mmsADORst = Nothing
LocalError:
    Exit Sub
End Sub
Private Sub DeleteTemporary()
On Error GoTo LocalError
    mmsAdoCmd.CommandText = "Delete From PRSDetailsTemp"
    Set mmsADORst = mmsAdoCmd.Execute
LocalError:
    Exit Sub
End Sub

Private Function GetNextPRSID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(PRSID) AS MaxID FROM PRSDetails"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextPRSID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextPRSID = 1
       Else
           GetNextPRSID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtPRSDate.Text = "" Then
        MsgBox "Fill-up PRS Date", vbExclamation, "PRS Date Required"
        txtPRSDate.SetFocus
        Exit Function
    End If
    If txtPRSNeed.Text = "" Then
        MsgBox "Fill-up PRS Date Needed", vbExclamation, "PRS Date Required"
        txtPRSNeed.SetFocus
        Exit Function
    End If
    If txtPRSRemark.Text = "" Then
        MsgBox "Fill-up PRS Remark", vbExclamation, "PRS Remark Required"
        txtPRSRemark.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function DataItemValidation() As Boolean
On Error GoTo LocalError
  DataItemValidation = False
    If txtPRSItem.Text = "" Then
        MsgBox "No Item Name.", vbExclamation, "Item Required"
        txtPRSItem.SetFocus
        Exit Function
    End If
    If txtPRSQty.Text = "" Or Val(txtPRSQty.Text) = 0 Then
        MsgBox "No Quantity of Item.", vbExclamation, "Quantity Required"
        txtPRSQty.SetFocus
        Exit Function
    End If
    If txtPRSUnit.Text = "" Then
        MsgBox "No Item Unit.", vbExclamation, "Item Unit Required"
        txtPRSUnit.SetFocus
        Exit Function
    End If
    DataItemValidation = True
LocalError:
    'MsgBox "Check Data", vbExclamation, "Check"
    Exit Function
End Function

Private Sub FormatText()
   txtPRSQty.Text = Format$(txtPRSQty, "#,###.#0")
   txtPRSBudget.Text = Format$(txtPRSBudget, "#,###.#0")
   txtPRSServed.Text = Format$(txtPRSServed, "#,###.#0")
   txtPRSInv.Text = Format$(txtPRSInv, "#,###.#0")
End Sub
Private Sub ClearBox()
    txtPRSNum.Text = ""
    txtPRSDate.Text = "__/__/____"
    txtPRSNeed.Text = "__/__/____"
    txtPRSRemark.Text = ""
    lvwPRS.ListItems.Clear
End Sub
Private Sub ClearItemBox()
    txtPRSItem.Text = ""
    ItemCode = ""
    ItemGroup = ""
    txtPRSItemCode.Text = ""
    txtPRSQty.Text = ""
    txtPRSUnit.Text = ""
    txtPRSDept.Clear
    txtPRSTransact.Clear
    txtPRSBudget.Text = ""
    txtPRSServed.Text = ""
    txtPRSInv.Text = ""
End Sub
Private Sub ClearFrame()
    frameItemAdd.Visible = False
    frameSearch.Visible = False
    frameItems.Visible = False
End Sub
Private Sub ItemBoxState(boxEnabled As Boolean)
    txtPRSItem.Enabled = boxEnabled
    txtPRSQty.Enabled = boxEnabled
    txtPRSUnit.Enabled = boxEnabled
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtPRSDate.Enabled = boxEnabled
    txtPRSNeed.Enabled = boxEnabled
    txtPRSFarm.Enabled = boxEnabled
    txtPRSRemark.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwPRS.Enabled = buttonEnabled
    cmdNew.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdSearch.Enabled = buttonEnabled
    cmdPrint.Enabled = buttonEnabled
    'cmdEdit.Enabled = buttonEnabled
    'cmdDelete.Enabled = buttonEnabled
End Sub
Private Sub Charging()
    txtPRSDept.Enabled = True
    txtPRSDept.SetFocus
    txtPRSDept.Text = ""
    txtPRSRemark.Text = ""
End Sub
Private Sub SendKeys_()
    'SendKeys "{left}"
    'SendKeys "{del}"
End Sub
'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HC0FFFF
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFFF
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdEdit_GotFocus()
   cmdEdit.BackColor = &HC0FFFF
End Sub
Private Sub cmdEdit_LostFocus()
   cmdEdit.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0FFFF
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub
Private Sub cmdSave_GotFocus()
   'cmdSave.BackColor = &HC0FFFF
End Sub
Private Sub cmdSave_LostFocus()
   'cmdSave.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFFF
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HC0FFFF
End Sub
Private Sub cmdSearch_LostFocus()
   cmdSearch.BackColor = &H8000000F
End Sub
Private Sub cmdPrint_GotFocus()
   cmdPrint.BackColor = &HC0FFFF
End Sub
Private Sub cmdPrint_LostFocus()
   cmdPrint.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFFF
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub cmdSavePRSDetails_GotFocus()
   cmdSavePRSDetails.BackColor = &HC0FFFF
End Sub
Private Sub cmdSavePRSDetails_LostFocus()
   cmdSavePRSDetails.BackColor = &H8000000F
End Sub

