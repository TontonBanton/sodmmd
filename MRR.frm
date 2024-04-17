VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMRR 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MATERIALS RECEIVING"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MRR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   36
      Top             =   9720
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
         Left            =   4730
         Style           =   1  'Graphical
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   4
         Top             =   150
         Width           =   2300
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
         TabIndex        =   3
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
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   2235
      End
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
      Height          =   550
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   1550
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
      Height          =   550
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   1550
   End
   Begin VB.Frame frameItemAdd 
      BackColor       =   &H00FFFFC0&
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
      Left            =   240
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   7045
      Begin VB.ComboBox txtMRSRef 
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
         ItemData        =   "MRR.frx":1E72
         Left            =   2000
         List            =   "MRR.frx":1E74
         TabIndex        =   56
         Text            =   "txtMRSRef"
         Top             =   2520
         Width           =   4650
      End
      Begin VB.TextBox txtMRSReq 
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
         Height          =   615
         Left            =   2000
         TabIndex        =   40
         Top             =   3120
         Width           =   4650
      End
      Begin VB.TextBox txtItemAmount 
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
         Left            =   2000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Top             =   5280
         Width           =   4650
      End
      Begin VB.CommandButton cmdSaveMRDetails 
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6300
         Width           =   4560
      End
      Begin VB.TextBox txtMRSQty 
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
         Left            =   2000
         MaxLength       =   50
         TabIndex        =   39
         Top             =   4080
         Width           =   4650
      End
      Begin VB.TextBox txtMRSItem 
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
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         Top             =   1050
         Width           =   6300
      End
      Begin VB.TextBox txtItemCost 
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
         Left            =   2000
         MaxLength       =   50
         TabIndex        =   43
         Top             =   4680
         Width           =   4650
      End
      Begin VB.TextBox txtMRSUnit 
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
         Height          =   600
         Left            =   2000
         MaxLength       =   50
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4650
      End
      Begin VB.TextBox txtMRSGroup 
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
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   400
         Width           =   6300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "AREA/PH"
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
         TabIndex        =   57
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Index           =   1
         Left            =   360
         TabIndex        =   55
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   350
         TabIndex        =   53
         Top             =   5445
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Index           =   10
         Left            =   360
         TabIndex        =   52
         Top             =   3240
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "RECEIVED"
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
         TabIndex        =   51
         Top             =   4200
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   350
         TabIndex        =   42
         Top             =   4800
         Width           =   705
      End
   End
   Begin VB.Frame frameItems 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1455
      Left            =   11280
      TabIndex        =   34
      Top             =   4560
      Visible         =   0   'False
      Width           =   2385
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
         Left            =   13000
         TabIndex        =   18
         Top             =   240
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
         Left            =   9500
         TabIndex        =   17
         Top             =   240
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
         Left            =   240
         MaxLength       =   50
         TabIndex        =   15
         Top             =   240
         Width           =   9255
      End
      Begin MSComctlLib.ListView lvwMRSItems 
         Height          =   5460
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   9631
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
   Begin VB.Frame frameSuppliers 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1545
      Left            =   11400
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton cmdAddSupplier 
         Caption         =   "SUPPLIER LIBRARY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   14
         Top             =   7200
         Width           =   9480
      End
      Begin VB.TextBox txtSupSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         MaxLength       =   50
         TabIndex        =   12
         Top             =   285
         Width           =   9480
      End
      Begin MSComctlLib.ListView lvwMRSSuppliers 
         Height          =   6210
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   10954
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
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   28
      Top             =   750
      Width           =   16935
      Begin VB.TextBox txtMRSRemark 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7200
         TabIndex        =   11
         Top             =   1050
         Width           =   9400
      End
      Begin VB.TextBox txtMRSId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSMask.MaskEdBox txtMRSDate 
         Height          =   420
         Left            =   2040
         TabIndex        =   6
         Top             =   645
         Width           =   2700
         _ExtentX        =   4763
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
      Begin VB.TextBox txtMRSRef1 
         Enabled         =   0   'False
         Height          =   420
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMRSSupplier 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7200
         TabIndex        =   9
         Top             =   645
         Width           =   9400
      End
      Begin VB.TextBox txtMRSPrs 
         Enabled         =   0   'False
         Height          =   420
         Left            =   2040
         TabIndex        =   8
         Top             =   1050
         Width           =   2700
      End
      Begin VB.TextBox txtMRSPoPcv 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7200
         TabIndex        =   7
         Top             =   240
         Width           =   9400
      End
      Begin VB.TextBox txtMRSNum 
         Enabled         =   0   'False
         Height          =   420
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   5
         Top             =   240
         Width           =   2700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   6000
         TabIndex        =   48
         Top             =   1100
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "SUPPLIER"
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
         Left            =   6000
         TabIndex        =   22
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "REF / PO"
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
         Index           =   6
         Left            =   6000
         TabIndex        =   23
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Index           =   5
         Left            =   360
         TabIndex        =   24
         Top             =   1100
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "MRS DATE"
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
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "MRS NUMBER"
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
         TabIndex        =   26
         Top             =   300
         Width           =   1290
      End
   End
   Begin VB.TextBox txtMRSTotal 
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
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   54
      Top             =   9720
      Width           =   3540
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
      Height          =   1290
      Left            =   11280
      TabIndex        =   31
      Top             =   6120
      Visible         =   0   'False
      Width           =   3075
      Begin VB.TextBox txtMRSSearch 
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
         Left            =   3250
         MaxLength       =   50
         TabIndex        =   20
         Top             =   240
         Width           =   5000
      End
      Begin VB.ComboBox cboMRSSearch 
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
         ItemData        =   "MRR.frx":1E76
         Left            =   150
         List            =   "MRR.frx":1E89
         TabIndex        =   19
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.ListView lvwMRSSearch 
         Height          =   1020
         Left            =   0
         TabIndex        =   21
         Top             =   870
         Width           =   16875
         _ExtentX        =   29766
         _ExtentY        =   1799
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
   Begin MSComctlLib.ListView lvwMRS 
      Height          =   7140
      Left            =   0
      TabIndex        =   30
      Top             =   2400
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   12594
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
      BackColor       =   &H00FFFFC0&
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
      Left            =   12120
      TabIndex        =   35
      Top             =   9960
      Width           =   1050
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   120
      Picture         =   "MRR.frx":1EAF
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2325
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
      Left            =   6844
      TabIndex        =   47
      Top             =   150
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
      Left            =   6844
      TabIndex        =   46
      Top             =   380
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   5938
      Picture         =   "MRR.frx":2A10
      Stretch         =   -1  'True
      Top             =   70
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   16935
   End
End
Attribute VB_Name = "FormMRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String

Private MRSId, EncodeMode, ItemName, ItemCode, ItemGroup, MRSResult, MRSInfo    As String
Private ItemCodeDel, ItemNoDel, ItemQtyDel, TxtVal, NumVal, AvailStock          As String
Private MRSTotal, MRSAvail, MRSDisc, ItemAmount, MRSRef                         As Double
Private EndCost, EndAmount, EndQty, MRRItemNo                                   As Double
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwMRS
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
    lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdNew_Click()
Dim MRSGroup As String
    EncodeMode = "A"
    ButtonState False
    cmdAdd.Enabled = True
    cmdCancel.Enabled = True
    BoxState True
    ClearBox
    ClearItemBox
    ClearFrame
    lvwMRS.ListItems.Clear
     MRSGroup = FormMainMenu.lblArea.Caption & "-" & "MRS"
     MRSId = Format$(GetNextMRSID, "000000")
     MRRItemNo = 0
     txtMRSDate = Format$(Now, "mm/dd/yyyy")
     txtMRSNum.Text = MRSGroup & "-" & MRSId
     txtMRSDate.SetFocus
     DeleteTemporary
            strsql = " Insert Into InventoryTemp Select * From Inventory"
            CommandExecute
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ClearFrame
    ClearItemBox
    LoadItemsList
End Sub
Private Sub cmdSaveMRDetails_Click()
   If Not DataItemValidation Then
        Exit Sub
   End If
   MRRItemNo = MRRItemNo + 1
   GetItemAmount
   SaveTemporary
   SaveInventoryTemp
       'SaveStockCard
   LoadMRSDetails
    
   ClearFrame
   cmdPrint.Enabled = True
   lvwMRS.Enabled = True
   
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
    lvwMRS.ListItems.Clear
    cmdNew.SetFocus
End Sub
Private Sub cmdEdit_Click()
   'MsgBox "Edit"
End Sub
Private Sub cmdDelete_Click()
    'mmsAdoCmd.CommandText = "Delete From MRSDetails"
    'Set mmsADORst = mmsAdoCmd.Execute
    'mmsAdoCmd.CommandText = "Update Materials SET Cost = '" & 0 & "'"
    'Set mmsADORst = mmsAdoCmd.Execute
End Sub
Private Sub cmdSearch_Click()
   EncodeMode = "S"
   ClearBox
   ClearItemBox
   ClearFrame
   DeleteTemporary
   BoxState False
   LoadSearchList
   txtMRSSearch.Text = ""
   cboMRSSearch.SetFocus
End Sub
Private Sub cmdPrint_Click()
    'If MsgBox("Auto-issuance received items?", _
    '    vbYesNo + vbQuestion, "Exit") = vbYes Then
    '     AutoIssuance
    'End If
    
    If EncodeMode = "A" Or EncodeMode = "E" Then
      strsql = " Insert Into MRSDetails Select * From MRSDetailsTemp "
      CommandExecute
      strsql = "Delete From Inventory": CommandExecute
      strsql = " Insert Into Inventory Select * From InventoryTemp "
      CommandExecute
      'strsql = " Insert Into StockCard Select * From StockCardTemp "
      'CommandExecute
    End If
    
    ClearBox
    ClearItemBox
    ClearFrame
    BoxState False
    ButtonState False
    cmdNew.Enabled = True
    cmdSearch.Enabled = True
    
   Load DataEnvironment1
   If DataEnvironment1.rsCommand1.State <> 0 Then DataEnvironment1.rsCommand1.Close
     ReportMRS.Refresh
   If ReportMRS.Visible = False Then ReportMRS.Show
   
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
Private Sub cmdAddSupplier_Click()
       FormSuppliers.Show
End Sub
Private Sub cmdAddItem_Click()
       FormItems.Show
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



'---------------------------------------------------------------------------------
'                                   T E X T B O X   E V E N T S
'---------------------------------------------------------------------------------
Private Sub TxtMRSdate_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      'SendKeys_
   ElseIf KeyAscii = 13 Then
      txtMRSPrs.SetFocus
   End If
End Sub
Private Sub txtMRSDate_GotFocus()
   txtMRSDate.SelLength = Len(txtMRSDate.Text)
End Sub
Private Sub txtMRSPRS_GotFocus()
    txtMRSPrs.Text = FormSettings.txtArea & "-" & "PRS" & "-"
    txtMRSPrs.SelStart = 8
    txtMRSPrs.SelLength = Len(txtMRSPrs.Text)
End Sub
Private Sub txtMRSPRS_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      txtMRSPoPcv.SetFocus
    End If
End Sub
Private Sub txtMRSPoPcv_GotFocus()
  If Not IsDate(txtMRSDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtMRSDate.SetFocus
        txtMRSDate = Format$(Now, "mm/dd/yyyy")
  End If
   txtMRSPoPcv.Text = "PO" & "#"
   txtMRSPoPcv.SelStart = 4
   txtMRSPoPcv.SelLength = Len(txtMRSPoPcv.Text)
End Sub
Private Sub txtMRSPoPcv_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      txtMRSSupplier.SetFocus
    End If
End Sub
Private Sub txtMRSSupplier_GotFocus()
   txtMRSSupplier.Text = ""
   txtMRSSupplier.SelLength = Len(txtMRSSupplier.Text)
End Sub
Private Sub txtMRSSupplier_DblClick()
    frameSuppliers.Top = 750: frameSuppliers.Left = 10375: frameSuppliers.Width = 14850: frameSuppliers.Height = 7900
    frameSuppliers.Visible = True: BoxState False
    txtSupSearch.SetFocus
    With lvwMRSSuppliers
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "Name", .Width * 0.93
     End With
   LoadSuppliers
End Sub
Private Sub txtMRSSupplier_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii > 0 Then
      txtMRSSupplier_DblClick
    End If
End Sub
Private Sub txtMRSRemark_GotFocus()
   txtMRSRemark.SelLength = Len(txtMRSRemark.Text)
End Sub
Private Sub txtMRSRemark_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       If txtMRSRemark.Text = "" Then
          txtMRSRemark.Text = "-"
       End If
      LoadItemsList
    End If
   If Not IsDate(txtMRSDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtMRSDate.SetFocus
        txtMRSDate = Format$(Now, "mm/dd/yyyy")
  End If
End Sub
'--------

'--------
Private Sub txtMRSUnit_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
       txtMRSRef.SetFocus
   End If
End Sub
Private Sub txtMRSRef_GotFocus()
    strsql = "SELECT * FROM Area ORDER BY AreaID"
    CommandExecute
    With mmsADORst
    txtMRSRef.Clear
    Do While Not .EOF
        txtMRSRef.AddItem ![AreaName]: .MoveNext
    Loop
    End With
End Sub
Private Sub txtMRSRef_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       If txtMRSRef.Text = "" Then
          txtMRSRef.Text = "-"
       End If
       txtMRSReq.SetFocus
   End If
End Sub
Private Sub txtMRSReq_GotFocus()
 txtMRSReq.SelLength = Len(txtMRSReq.Text)
End Sub
Private Sub txtMRSReq_Change()
  TxtVal = txtMRSReq.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtMRSReq.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtMRSReq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtMRSQty.Text = txtMRSReq.Text
       txtMRSQty.SetFocus
   End If
End Sub

Private Sub txtMRSQty_GotFocus()
 txtMRSQty.SelLength = Len(txtMRSQty.Text)
End Sub
Private Sub txtMRSQty_Change()
  TxtVal = txtMRSQty.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtMRSQty.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtMRSQty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       'txtMRSReq.Text = txtMRSQty.Text
       txtItemCost.SetFocus
   End If
End Sub
Private Sub txtItemCost_GotFocus()
   txtItemCost.Text = "0.00"
   txtItemCost.SelLength = Len(txtItemCost.Text)
End Sub
Private Sub txtitemCost_Change()
  TxtVal = txtItemCost.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtItemCost.Text = "" 'CStr(NumVal)
  End If
End Sub
Private Sub txtitemCost_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not DataItemValidation Then
        Exit Sub
      End If
         GetItemAmount
         cmdSaveMRDetails.SetFocus
   End If
End Sub
Private Sub GetItemAmount()
        txtMRSQty.Text = Format$(txtMRSQty, "#,###.#0")
        ItemAmount = CDbl(txtItemCost.Text) * CDbl(txtMRSQty.Text)
        txtItemAmount.Text = Format$(ItemAmount, "#,###.#0")
        txtItemCost.Text = Format$(txtItemCost, "#,###.#0")
        GetDiscrepancy
End Sub
Private Sub GetDiscrepancy()
         If CDbl(txtMRSReq.Text) > CDbl(txtMRSQty.Text) Then
            MRSDisc = CDbl(txtMRSReq.Text) - CDbl(txtMRSQty.Text)
         ElseIf CDbl(txtMRSReq.Text) < CDbl(txtMRSQty.Text) Then
            MRSDisc = CDbl(txtMRSQty.Text) - CDbl(txtMRSReq.Text)
         End If
         If MRSDisc = 0 Then
            MRSDisc = CStr("-")
         End If
         
         MRSDisc = Format$(MRSDisc, "#,###.#0")
End Sub
'------------------------------------------------------------------------------------
'                      S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboMRSSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtMRSSearch.SetFocus
   End If
   If KeyAscii < 255 Then
      'SendKeys_
   End If
End Sub
Private Sub cboMRSSearch_GotFocus()
  cboMRSSearch.Text = "NUMBER"
End Sub
Private Sub txtMRSSearch_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     lvwMRSSearch.SetFocus
  End If
End Sub
Private Sub txtMRSSearch_Change()
Dim MRSSearchLI       As ListItem
On Error GoTo LocalError
    If cboMRSSearch.Text = "NUMBER" Then
       strsql = "Select * from MRSDetails where MRSID like '" & txtMRSSearch.Text & "%'" & "Order by MRSNum"
    ElseIf cboMRSSearch.Text = "DATE" Then
       strsql = "Select * from MRSDetails where MRSDate like '" & txtMRSSearch.Text & "%'" & "Order by MRSNum"
    ElseIf cboMRSSearch.Text = "PO" Then
       strsql = "Select * from MRSDetails where MRSPoPcv like '" & txtMRSSearch.Text & "%'" & "Order by MRSNum"
    ElseIf cboMRSSearch.Text = "ITEM" Then
       strsql = "Select * from MRSDetails where MRSItem like '" & txtMRSSearch.Text & "%'" & "Order by MRSNum"
    ElseIf cboMRSSearch.Text = "SUPPLIER" Then
       strsql = "Select * from MRSDetails where MRSSupplier like '" & txtMRSSearch.Text & "%'" & "Order by MRSNum"
    End If
 
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LoadMRSSearch

LocalError:
    Exit Sub
End Sub
'---------- S E A R C H   L I S T V I E W-----
Private Sub lvwMRSSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
       SearchLoad
    End With
End Sub
Private Sub lvwMRSSearch_DblClick()
   frameSearch.Visible = False
   BoxState False
   SaveTemporary
   LoadMRSDetails
   cmdPrint.Enabled = True
End Sub
Private Sub lvwMRSSearch_KeyPress(KeyAscii As Integer)
On Error GoTo LocalError
   If KeyAscii = 13 Then
       lvwMRSSearch_DblClick
       MRSResult = lvwMRSSearch.SelectedItem.Text
   End If
LocalError:
    Exit Sub
End Sub
Private Sub lvwMRSSearch_GotFocus()
On Error GoTo LocalError
    SearchLoad
LocalError:
    Exit Sub
End Sub
Private Sub SearchLoad()
      txtMRSNum.Text = lvwMRSSearch.SelectedItem.Text
      txtMRSDate.Text = Format$(lvwMRSSearch.SelectedItem.SubItems(1), "mm/dd/yyyy")
      txtMRSPoPcv.Text = lvwMRSSearch.SelectedItem.SubItems(2)
      txtMRSPrs.Text = lvwMRSSearch.SelectedItem.SubItems(3)
      txtMRSSupplier.Text = lvwMRSSearch.SelectedItem.SubItems(5)
      txtMRSRemark.Text = lvwMRSSearch.SelectedItem.SubItems(6)
End Sub
Private Sub LoadSearchList()
    frameSearch.Top = lvwMRS.Top - 100: frameSearch.Left = lvwMRS.Left: frameSearch.Height = lvwMRS.Height + 100: frameSearch.Width = lvwMRS.Width
    lvwMRSSearch.Height = frameSearch.Height - 900
    frameSearch.Visible = True

    lvwMRSItems.ColumnHeaders.Clear
    With lvwMRSSearch
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NUMBER", .Width * 0.12
        .ColumnHeaders.Add , , "DATE", .Width * 0.07
        .ColumnHeaders.Add , , "PO/PCV #", .Width * 0.08
        .ColumnHeaders.Add , , "PRS", .Width * 0#
        .ColumnHeaders.Add , , "ITEM", .Width * 0.25
        .ColumnHeaders.Add , , "SUPPLIER", .Width * 0.2
        .ColumnHeaders.Add , , "REMARK", .Width * 0.17
    End With

    strsql = "SELECT DISTINCT MRSNum, MRSDate, MRSPoPcv, MRSPrs" _
           & "     , MRSItem, MRSSupplier, MRSRemark" _
           & "     FROM MRSDetails ORDER BY MRSNum "
    CommandExecute
    
    LoadMRSSearch
End Sub
Private Sub LoadMRSSearch()
Dim MRSSearchLI   As ListItem
On Error GoTo LocalError

    lvwMRSSearch.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MRSSearchLI = lvwMRSSearch.ListItems.Add(, , !MRSNum & "")
            MRSSearchLI.SubItems(1) = !MRSDate & ""
            MRSSearchLI.SubItems(2) = !MRSPoPcv & ""
            MRSSearchLI.SubItems(3) = !MRSPrs & ""
            MRSSearchLI.SubItems(4) = !MRSItem & ""
            MRSSearchLI.SubItems(5) = !MRSSupplier & ""
            MRSSearchLI.SubItems(6) = !MRSRemark & ""
            .MoveNext
        Loop
    End With
     
    Set MRSSearchLI = Nothing
    Set mmsADORst = Nothing
    
LocalError:
    Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                  I T E M S   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtItemSearch_Change()
Dim MaterialsLI  As ListItem
    strsql = "Select * from Materials where ItemName like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemName"
             
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwMRSItems.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwMRSItems.ListItems.Add(, , !ItemGroup & "")
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
Private Sub txtItemSearch_Click()
    txtItemSearch.Text = ""
End Sub
Private Sub txtItemSearch_GotFocus()
   txtItemSearch.Text = ""
   txtItemSearch.SelLength = Len(txtItemSearch.Text)
End Sub
Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   'If KeyCode = 112 Then
   '    cmdAddItem_Click
   'End If
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      lvwMRSItems.SetFocus
   End If
End Sub
'---------- I T E M S  L I S T V I E W ----------
Private Sub lvwMRSItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      ItemGroup = lvwMRSItems.SelectedItem.Text
      ItemCode = lvwMRSItems.SelectedItem.SubItems(1)
      LoadItemSearch
    End With
End Sub
Private Sub lvwMRSItems_DblClick()
    frameItems.Visible = False
    frameItemAdd.Top = lvwMRS.Top - 100: frameItemAdd.Left = 7500: frameItemAdd.Height = 7500: frameItemAdd.Width = 10500
    BoxState False
    ItemBoxState True
    'CheckInventory
    frameItemAdd.Visible = True
    txtMRSRef.SetFocus
End Sub
Private Sub CheckInventory()
On Error GoTo LocalError
    'strsql = "Select * From InventoryEndingTemp where ItemCode like '" & ItemCode & "' "
    'mmsAdoCmd.CommandText = strsql
    'Set mmsADORst = mmsAdoCmd.Execute
    
    If mmsADORst.EOF Or mmsADORst.BOF Then
       'MsgBox "new in SaveInv"
       strsql = "Select * From Materials where ItemCode like '" & ItemCode & "' "
       mmsAdoCmd.CommandText = strsql
       Set mmsADORst = mmsAdoCmd.Execute
       
       EndQty = mmsADORst.Fields("AvailStock")
       EndAmount = mmsADORst.Fields("ItemAmount")
       EndCost = mmsADORst.Fields("ItemCost")
 
    Else
       'MsgBox "already in SaveInv"
       'strsql = "Select * From InventoryEndingTemp where ItemCode like '" & ItemCode & "' "
       'mmsAdoCmd.CommandText = strsql
       'Set mmsADORst = mmsAdoCmd.Execute
       
       EndQty = mmsADORst.Fields("ItemQty")
       EndAmount = mmsADORst.Fields("ItemAmount")
       EndCost = mmsADORst.Fields("ItemCost")
    End If
LocalError:
   Exit Sub
End Sub
Private Sub lvwMRSItems_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
      cmdAddItem_Click
   End If
End Sub
Private Sub lvwMRSItems_KeyPress(KeyAscii As Integer)
   lvwMRSItems_DblClick
End Sub
Private Sub lvwMRSItems_GotFocus()
On Error GoTo LocalError
      ItemGroup = lvwMRSItems.SelectedItem.Text
      ItemCode = lvwMRSItems.SelectedItem.SubItems(1)
      LoadItemSearch
LocalError:
    Exit Sub
End Sub
Private Sub LoadItemSearch()
      txtMRSGroup.Text = ItemGroup & "-" & ItemCode
      txtMRSItem.Text = lvwMRSItems.SelectedItem.SubItems(2)
      txtMRSUnit.Text = lvwMRSItems.SelectedItem.SubItems(4)
End Sub
Private Sub LoadItemsList()
Dim MaterialsLI  As ListItem
    
    If Not DataValidation Then
       Exit Sub
    End If
    
       frameItems.Top = Frame1.Top: frameItems.Left = lvwMRS.Left: frameItems.Width = lvwMRS.Width: frameItems.Height = 8900
       lvwMRSItems.Top = 950: lvwMRSItems.Height = frameItems.Height - 800: frameItems.Visible = True: txtItemSearch.SetFocus
    
     lvwMRSItems.ColumnHeaders.Clear
     With lvwMRSItems
        .ColumnHeaders.Add , , "GROUP", .Width * 0#
        .ColumnHeaders.Add , , "CODE", .Width * 0.1
        .ColumnHeaders.Add , , "ITEM NAME", .Width * 0.6
        .ColumnHeaders.Add , , "STOCK", .Width * 0#
        .ColumnHeaders.Add , , " ", .Width * 0.08
        .ColumnHeaders.Add , , "ID", Width * 0#
     End With
     lvwMRSItems.ColumnHeaders.Item(4).Alignment = lvwColumnRight
    
    strsql = "SELECT * From Materials Order By ItemName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwMRSItems.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwMRSItems.ListItems.Add(, , !ItemGroup & "")
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
'------------------------------------------------------------------------------------
'                   S U P P L I E R   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtSupSearch_Change()
Dim SuppliersLI  As ListItem
    
    strsql = "Select * from Suppliers where SupName like '" & txtSupSearch.Text & "%'" _
             & "Order by SupName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute

    lvwMRSSuppliers.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set SuppliersLI = lvwMRSSuppliers.ListItems.Add(, , !SupName & "")
           .MoveNext
        Loop
     End With
    Set SuppliersLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub txtSupSearch_Click()
    txtSupSearch.Text = ""
End Sub
Private Sub txtSupSearch_GotFocus()
   txtSupSearch.Text = ""
   txtSupSearch.SelLength = Len(txtSupSearch.Text)
End Sub
Private Sub txtSupSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
      cmdAddSupplier_Click
   End If
End Sub
Private Sub txtSupSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwMRSSuppliers.SetFocus
   End If
End Sub
'---------- S U P P L I E R S  L I S T V I E W -----
Private Sub lvwMRSSuppliers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtMRSSupplier.Text = lvwMRSSuppliers.SelectedItem.Text
    End With
End Sub
Private Sub lvwMRSSuppliers_DblClick()
    frameSuppliers.Visible = False
    BoxState True
    txtMRSRemark.SetFocus
End Sub
Private Sub lvwMRSSuppliers_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddSupplier_Click
   End If
End Sub
Private Sub lvwMRSSuppliers_KeyPress(KeyAscii As Integer)
    lvwMRSSuppliers_DblClick
End Sub
Private Sub lvwMRSSuppliers_GotFocus()
On Error GoTo LocalError
      txtMRSSupplier.Text = lvwMRSSuppliers.SelectedItem.Text
LocalError:
    Exit Sub
End Sub
'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub lvwMRS_DblClick()
   If lvwMRS.SelectedItem Is Nothing Then
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
    ItemNoDel = lvwMRS.SelectedItem
    ItemCodeDel = lvwMRS.SelectedItem.SubItems(1)
    ItemQtyDel = lvwMRS.SelectedItem.SubItems(3)
    
        '-----
        strsql = "Delete From MRSDetailsTemp where MRSItemNo like " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
    
        strsql = "Update MRSDetailsTemp SET MRSItemNo = MRSItemNo - 1"
        strsql = strsql & " where MRSItemNo > " & ItemNoDel & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute

        '------
        'strsql = "Delete From StockCardTemp where StockID like " & ItemNoDel & ""
        'mmsAdoCmd.CommandText = strsql
        'Set mmsADORst = mmsAdoCmd.Execute
        
        '------
        'strsql = "Select * From InventoryEndingTemp where ItemCode like '" & ItemCodeDel & "'"
        'mmsAdoCmd.CommandText = strsql
        'Set mmsADORst = mmsAdoCmd.Execute
        'EndQty = mmsADORst.Fields("ItemQty") - ItemQtyDel
        'EndCost = mmsADORst.Fields("ItemCost")
        'EndAmount = Format$(EndQty * EndCost, "#,###.#0")
        'UpdateInventoryTemp
        
        MRRItemNo = MRRItemNo - 1
          
    LoadMRSDetails
    
End Sub
' --------------  S E T   L I S T V I E W  -----------
Private Sub SetlvwMRS()
lvwMRS.ColumnHeaders.Clear
    With lvwMRS
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " # ", .Width * 0.03
        .ColumnHeaders.Add , , "CODE", .Width * 0.07
        .ColumnHeaders.Add , , "ITEM", .Width * 0.35
        .ColumnHeaders.Add , , "QTY", Width * 0.08
        .ColumnHeaders.Add , , "UNIT", .Width * 0.05
        .ColumnHeaders.Add , , "U/P", Width * 0.12
        .ColumnHeaders.Add , , "AMOUNT", Width * 0.12
        .ColumnHeaders.Add , , "REQD", Width * 0.08
        .ColumnHeaders.Add , , "DISC", Width * 0.08
    End With
  lvwMRS.ColumnHeaders.Item(4).Alignment = lvwColumnRight
  lvwMRS.ColumnHeaders.Item(6).Alignment = lvwColumnRight
  lvwMRS.ColumnHeaders.Item(7).Alignment = lvwColumnRight
  lvwMRS.ColumnHeaders.Item(8).Alignment = lvwColumnRight
  lvwMRS.ColumnHeaders.Item(9).Alignment = lvwColumnRight
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
Private Sub LoadMRSDetails()
Dim MRSLI             As ListItem
On Error GoTo LocalError

      GetMRSTotal
        txtMRSTotal.Text = Format$(MRSTotal, "#,###.#0")
      InsertMRSTotal
    
    strsql = "SELECT * from MRSDetailsTemp ORDER BY MRSItemNo "
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    lvwMRS.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set MRSLI = lvwMRS.ListItems.Add(, , !MRSItemNo & "")
            'MRSLI.SubItems(1) = !MRSCode & ""
            MRSLI.SubItems(2) = !MRSItem & ""
            MRSLI.SubItems(3) = !MRSQty & ""
            MRSLI.SubItems(4) = !MRSUnit & ""
            MRSLI.SubItems(5) = Format$(!MRSCost, "#,###.#0") & ""
            MRSLI.SubItems(6) = Format$(!MRSAmount, "#,###.#0") & ""
            MRSLI.SubItems(7) = !MRSReqd & ""
            MRSLI.SubItems(8) = !MRSDisc & ""
            .MoveNext
        Loop
      End With

    cmdAdd.SetFocus
LocalError:
    Exit Sub
End Sub
Private Sub GetMRSTotal()
        strsql = " SELECT SUM(MRSAmount) as SubTotal FROM MRSDetailsTemp "
        CommandExecute
        MRSTotal = mmsADORst.Fields!Subtotal
        Set mmsADORst = Nothing
        txtMRSTotal.Text = Format$(MRSTotal, "#,###.#0")
End Sub
Private Sub InsertMRSTotal()
On Error GoTo LocalError
        strsql = "Update MRSDetailsTemp SET MRSTotal = '" & txtMRSTotal.Text & "'"
        strsql = strsql & " where MRSNum like '" & txtMRSNum.Text & "'"
        CommandExecute
LocalError:
    Exit Sub
End Sub
Private Sub LoadSuppliers()
Dim SuppliersLI  As ListItem
On Error GoTo LocalError
                                 
strsql = "SELECT SupName from Suppliers" _
           & " ORDER BY SupName"

    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwMRSSuppliers.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set SuppliersLI = lvwMRSSuppliers.ListItems.Add(, , !SupName & "")
            .MoveNext
        Loop
     End With
    Set SuppliersLI = Nothing
    Set mmsADORst = Nothing
    
LocalError:
    Exit Sub
End Sub
Private Sub SaveTemporary()
Dim strsql          As String
On Error GoTo LocalError

  If EncodeMode = "A" Then
         strsql = "INSERT INTO MRSDetailsTemp (  MRSID"
         strsql = strsql & "            , MRSItemNo"
         strsql = strsql & "            , MRSNum"
         strsql = strsql & "            , MRSDate"
         strsql = strsql & "            , MRSPrs"
         strsql = strsql & "            , MRSPoPCv"
         strsql = strsql & "            , MRSSupplier"
         strsql = strsql & "            , MRSRemark"
         strsql = strsql & "            , MRSItem"
         strsql = strsql & "            , MRSGroup"
         strsql = strsql & "            , MRSCode"
         strsql = strsql & "            , MRSUnit"
         strsql = strsql & "            , MRSArea"
         strsql = strsql & "            , MRSQty"
         strsql = strsql & "            , MRSCost"
         strsql = strsql & "            , MRSAmount"
         strsql = strsql & "            , MRSReqd"
         strsql = strsql & "            , MRSDisc"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & MRSId
         strsql = strsql & ", '" & MRRItemNo & "'"
         strsql = strsql & ", '" & Replace$(txtMRSNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSPrs.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSPoPcv.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSSupplier.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSRemark.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemGroup & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtMRSUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSRef.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSQty.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemAmount.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSReq.Text, "'", "''") & "'"
         strsql = strsql & ", '" & MRSDisc & "'"
         strsql = strsql & ")"
         
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
  End If
  
  If EncodeMode = "S" Then
     strsql = " Insert Into MRSDetailsTemp Select * From MRSDetails Where MRSNum like '" & txtMRSNum.Text & "' Order By MRSItemNo "
     mmsAdoCmd.CommandText = strsql
     Set mmsADORst = mmsAdoCmd.Execute
     cmdCancel.Enabled = True
  End If
  
LocalError:
    Exit Sub
End Sub
Private Sub SaveInventoryTemp()
On Error GoTo LocalError
Dim lngIDField As Integer

  lngIDField = GetNextInvID
  MRSInfo = txtMRSNum.Text & " - " & txtMRSDate.Text
  
  If EncodeMode = "A" Then
         strsql = "INSERT INTO InventoryTemp (  InvID"
         strsql = strsql & "            , ItemGroup"
         strsql = strsql & "            , ItemCode"
         strsql = strsql & "            , ItemName"
         strsql = strsql & "            , Unit"
         strsql = strsql & "            , Cost"
         strsql = strsql & "            , AvailStock"
         strsql = strsql & "            , ItemAmount"
         strsql = strsql & "            , Remarks"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & ItemGroup & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtMRSItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSQty.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemAmount.Text, "'", "''") & "'"
         strsql = strsql & ", '" & MRSInfo & "'"
         strsql = strsql & ")"
         
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
  End If
  
  If EncodeMode = "S" Then
     strsql = " Insert Into MRSDetailsTemp Select * From MRSDetails Where MRSNum like '" & txtMRSNum.Text & "' Order By MRSItemNo "
     mmsAdoCmd.CommandText = strsql
     Set mmsADORst = mmsAdoCmd.Execute
     cmdCancel.Enabled = True
  End If
LocalError:
   Exit Sub
End Sub
Private Sub SaveStockCard()
Dim strsql          As String
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
         strsql = strsql & MRRItemNo
         strsql = strsql & ", '" & Replace$(txtMRSDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemCode & "'"
         strsql = strsql & ", '" & Replace$(txtMRSItem.Text, "'", "''") & "'"
         strsql = strsql & ", '" & ItemGroup & "'"
             strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSRef.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtItemAmount.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtMRSQty.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
             strsql = strsql & ", '" & Replace$(CDbl(0#), "'", "''") & "'"
         strsql = strsql & ")"
         
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
  
LocalError:
    Exit Sub
End Sub
Private Sub AutoIssuance()
   MsgBox "save to mrdetails"
End Sub
Private Function GetNextMRSID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(MRSID) AS MaxID FROM MRSDetails"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextMRSID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextMRSID = 1
       Else
           GetNextMRSID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Function GetNextInvID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(InvID) AS MaxID FROM InventoryTemp"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextInvID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextInvID = 1
       Else
           GetNextInvID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtMRSDate.Text = "" Then
        MsgBox "Fill-up MRS Date.", vbExclamation, "MRS Date Required"
        txtMRSDate.SetFocus: Exit Function
    End If
    If txtMRSPoPcv.Text = "" Then
        MsgBox "Fill-up PO or PCV Number.", vbExclamation, "PO/PCV Required"
        txtMRSPoPcv.SetFocus: Exit Function
    End If
    If txtMRSPrs.Text = "" Then
        MsgBox "Fill-up MRS Number.", vbExclamation, "MRS Required"
        txtMRSPrs.SetFocus: Exit Function
    End If
    'If txtMRSRemark.Text = "" Then
    '    MsgBox "Fill-up Remarks for notation.", vbExclamation, "Reference Required"
    '    txtMRSRemark.SetFocus: Exit Function
    'End If
    If txtMRSSupplier.Text = "" Then
        MsgBox "Fill-up Supplier Name.", vbExclamation, "Supplier Required"
        txtMRSSupplier.SetFocus: Exit Function
    End If
    DataValidation = True

End Function
Private Function DataItemValidation() As Boolean
On Error GoTo LocalError
  DataItemValidation = False
    If txtMRSItem.Text = "" Or txtMRSUnit.Text = "" Then
        MsgBox "No Item Name.", vbExclamation, "Item Required"
        txtMRSItem.SetFocus: Exit Function
    End If
    If txtMRSQty.Text = "" Or CDbl(txtMRSQty.Text) = 0 Or txtMRSQty = Null Then
        MsgBox "No Quantity of Item.", vbExclamation, "Quantity Required"
        txtMRSQty.SetFocus: Exit Function
    End If

    DataItemValidation = True
LocalError:
    'MsgBox "Check Data", vbExclamation, "Check"
    Exit Function
End Function
Private Sub DeleteTemporary()
On Error GoTo LocalError
    strsql = "Delete From MRSDetailsTemp": CommandExecute
    strsql = "Delete From InventoryTemp": CommandExecute
    'mmsAdoCmd.CommandText = "Delete From StockCardTemp"
    'Set mmsADORst = mmsAdoCmd.Execute
    Set mmsADORst = Nothing
LocalError:
    Exit Sub
End Sub
Private Sub SendKeys_()
On Error GoTo LocalError
    SendKeys "{left}"
    SendKeys "{del}"
LocalError:
    Exit Sub
End Sub
Private Sub ClearBox()
    txtMRSNum.Text = ""
    txtMRSDate.Text = "__/__/____"
    txtMRSPoPcv.Text = ""
    txtMRSRemark.Text = ""
    txtMRSPrs.Text = ""
    txtMRSSupplier.Text = ""
    txtMRSRef.Clear
    txtMRSTotal.Text = ""
    lvwMRS.ListItems.Clear
End Sub
Private Sub ClearItemBox()
    txtMRSItem.Text = ""
    ItemCode = ""
    ItemGroup = ""
    txtMRSGroup.Text = ""
    txtMRSQty.Text = ""
    txtMRSReq.Text = ""
    txtMRSUnit.Text = ""
    txtItemCost.Text = ""
    txtItemAmount.Text = ""
    txtMRSTotal.Text = ""
    txtMRSRef.Clear
    MRSDisc = CStr("-")
End Sub
Private Sub ClearFrame()
    frameItemAdd.Visible = False
    frameSearch.Visible = False
    frameItems.Visible = False
    frameSuppliers.Visible = False
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtMRSDate.Enabled = boxEnabled
    txtMRSPoPcv.Enabled = boxEnabled
    txtMRSRemark.Enabled = boxEnabled
    txtMRSPrs.Enabled = boxEnabled
    txtMRSSupplier.Enabled = boxEnabled
End Sub
Private Sub ItemBoxState(boxEnabled As Boolean)
    txtMRSItem.Enabled = boxEnabled
    txtMRSQty.Enabled = boxEnabled
    txtMRSRef.Enabled = boxEnabled
    txtMRSUnit.Enabled = boxEnabled
    txtItemCost.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwMRS.Enabled = buttonEnabled
    cmdNew.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdSearch.Enabled = buttonEnabled
    cmdPrint.Enabled = buttonEnabled
    'cmdEdit.Enabled = buttonEnabled
    'cmdDelete.Enabled = buttonEnabled
End Sub


'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HFFFFC0
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HFFFFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdSaveMRDetails_GotFocus()
   cmdSaveMRDetails.BackColor = &HFFFFC0
End Sub
Private Sub cmdSaveMRDetails_LostFocus()
   cmdSaveMRDetails.BackColor = &H8000000F
End Sub
Private Sub cmdEdit_GotFocus()
   cmdEdit.BackColor = &HFFFFC0
End Sub
Private Sub cmdEdit_LostFocus()
   cmdEdit.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HFFFFC0
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub
Private Sub cmdSave_GotFocus()
   'cmdSave.BackColor = &HFFFFC0
End Sub
Private Sub cmdSave_LostFocus()
   'cmdSave.BackColor = &H8000000F
End Sub

Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HFFFFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HFFFFC0
End Sub
Private Sub cmdSearch_LostFocus()
   cmdSearch.BackColor = &H8000000F
End Sub
Private Sub cmdPrint_GotFocus()
   cmdPrint.BackColor = &HFFFFC0
End Sub
Private Sub cmdPrint_LostFocus()
   cmdPrint.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HFFFFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub



