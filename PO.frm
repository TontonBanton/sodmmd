VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormPO 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE ORDER"
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
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumWords 
      Enabled         =   0   'False
      Height          =   420
      Left            =   12600
      TabIndex        =   58
      Top             =   9000
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.TextBox txtPOSupAd 
      Enabled         =   0   'False
      Height          =   420
      Left            =   6960
      TabIndex        =   57
      Top             =   480
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   1785
      Left            =   9360
      TabIndex        =   33
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
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
         TabIndex        =   17
         Top             =   7920
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
         TabIndex        =   15
         Top             =   285
         Width           =   9480
      End
      Begin MSComctlLib.ListView lvwPOSuppliers 
         Height          =   7005
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   12356
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
      Height          =   550
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9000
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
      Height          =   550
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   9000
      Visible         =   0   'False
      Width           =   1550
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
      TabIndex        =   37
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
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.Frame frameItemAdd 
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
      Height          =   4860
      Left            =   840
      TabIndex        =   34
      Top             =   2880
      Visible         =   0   'False
      Width           =   7045
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
         TabIndex        =   44
         Top             =   3240
         Width           =   4650
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3960
         Width           =   4680
      End
      Begin VB.TextBox txtPOQty 
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
         TabIndex        =   40
         Top             =   2040
         Width           =   3210
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
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1050
         Width           =   6300
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
         Top             =   2640
         Width           =   4650
      End
      Begin VB.TextBox txtPOUnit 
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
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1410
      End
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
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   400
         Width           =   6300
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
         Left            =   360
         TabIndex        =   50
         Top             =   3480
         Width           =   1080
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
         Left            =   360
         TabIndex        =   49
         Top             =   2160
         Width           =   480
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
         Left            =   360
         TabIndex        =   42
         Top             =   2760
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
      Left            =   10080
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   4785
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   18
         Top             =   240
         Width           =   9255
      End
      Begin MSComctlLib.ListView lvwPOItems 
         Height          =   5460
         Left            =   0
         TabIndex        =   19
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
   Begin VB.Frame Frame1 
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
      Height          =   1695
      Left            =   -120
      TabIndex        =   30
      Top             =   600
      Width           =   17055
      Begin VB.ComboBox txtPOEquip 
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
         ItemData        =   "PO.frx":0000
         Left            =   13080
         List            =   "PO.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   645
         Width           =   3500
      End
      Begin VB.ComboBox txtPOTerms 
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
         ItemData        =   "PO.frx":0004
         Left            =   13080
         List            =   "PO.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   3500
      End
      Begin MSMask.MaskEdBox txtPODate 
         Height          =   420
         Left            =   2040
         TabIndex        =   6
         Top             =   645
         Width           =   2820
         _ExtentX        =   4974
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
      Begin VB.TextBox txtPOStatus 
         Enabled         =   0   'False
         Height          =   420
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "-"
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtPOSupplier 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   6000
      End
      Begin VB.TextBox txtPONum 
         Enabled         =   0   'False
         Height          =   420
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   5
         Top             =   240
         Width           =   2820
      End
      Begin VB.TextBox txtPOPrs 
         Enabled         =   0   'False
         Height          =   420
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   2805
      End
      Begin VB.TextBox txtPORemark 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7080
         TabIndex        =   14
         Top             =   1080
         Width           =   9500
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
         ItemData        =   "PO.frx":0038
         Left            =   2040
         List            =   "PO.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2820
      End
      Begin VB.TextBox txtPOWork2 
         Enabled         =   0   'False
         Height          =   420
         Left            =   7080
         TabIndex        =   55
         Top             =   645
         Width           =   6000
      End
      Begin VB.ComboBox txtPOWork 
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
         ItemData        =   "PO.frx":006C
         Left            =   7080
         List            =   "PO.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   645
         Width           =   6000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   5880
         TabIndex        =   48
         Top             =   1095
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   5880
         TabIndex        =   25
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DETAILS"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   27
         Top             =   1100
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PO DATE"
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
         TabIndex        =   28
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PO NUMBER"
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
         TabIndex        =   29
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.TextBox txtPOTotal 
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
      TabIndex        =   51
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
      Left            =   10080
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   4515
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
         Left            =   3250
         MaxLength       =   50
         TabIndex        =   23
         Top             =   240
         Width           =   5000
      End
      Begin VB.ComboBox cboPOSearch 
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
         ItemData        =   "PO.frx":0070
         Left            =   150
         List            =   "PO.frx":007A
         TabIndex        =   22
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.ListView lvwPOSearch 
         Height          =   1020
         Left            =   0
         TabIndex        =   24
         Top             =   840
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
   Begin MSComctlLib.ListView lvwPO 
      Height          =   7260
      Left            =   0
      TabIndex        =   31
      Top             =   2280
      Width           =   16935
      _ExtentX        =   29871
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
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   36
      Top             =   9960
      Width           =   1050
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
      Height          =   345
      Left            =   4680
      TabIndex        =   47
      Top             =   120
      Width           =   8325
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   4680
      TabIndex        =   46
      Top             =   360
      Width           =   8250
   End
End
Attribute VB_Name = "FormPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private POLI                   As ListItem

Private EncodeMode, ItemName, POResult, POSup                    As String
Private ItemCodeDel, ItemNoDel, ItemQtyDel, TxtVal, NumVal       As String
Private EndCost, EndAmount, EndQty, ItemAmount, PORef   As Double

Public PORemark, ItemCode, ItemGroup            As String
Public POId, POItemNo, SupId, POTotal           As Double

Private Sub Command1_Click()
'strsql = "Update PODetails SET POTr = '-'"
'CommandExecute
End Sub
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwPOMain
   lblComp.Caption = FormMainMenu.lblComp.Caption: lblHeader.Caption = FormMainMenu.lblHeader.Caption
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
  ButtonState False
  BoxState True
  ClearBox
  ClearFrame
  ClearItemBox
  DeleteTemporary
  EncodeMode = "A": cmdAdd.Enabled = True: cmdCancel.Enabled = True: lvwPO.ListItems.Clear: POItemNo = 0
    POId = Format$(GetNextPOID, "000000"): txtPONum.Text = POId
    txtPODate = Format$(Now, "mm/dd/yyyy"): txtPODate.SetFocus
    txtPOArea.Visible = True: txtPOPrs.Visible = False
 txtPOEquip_GotFocus
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ClearFrame
    ClearItemBox
    LoadItemsList
End Sub
Private Sub cmdSavePODetails_Click()
   If Not DataItemValidation Then
        Exit Sub
   End If
   GetItemAmount
   If EncodeMode = "A" Then
    POItemNo = POItemNo + 1
    SaveTemporary
   End If
   If EncodeMode = "E" Then
    strsql = "Update PODetailsTemp SET POQty = '" & txtPOQty.Text & "', POCost = '" & txtPOCost.Text & "'  where POItemNo like '" & POItemNo & "'":  CommandExecute
    ItemAmount = Format$(ItemAmount, "#,###.#0"): strsql = "Update PODetailsTemp SET POAmount = '" & ItemAmount & "' where POItemNo like '" & POItemNo & "'": CommandExecute
    EncodeMode = "A"
   End If
   
   LoadPODetails
   ClearFrame
   cmdPrint.Enabled = True: lvwPO.Enabled = True
End Sub
Private Sub cmdCancel_Click()
    If EncodeMode = "S" Then
        If MsgBox("Are you sure you want to cancel this PO?", vbYesNo + vbQuestion, "Exit") = vbYes Then
            strsql = "Update PODetails SET POStatus = 'CANCELLED', POWork = 'CANCELLED' WHERE POId = " & POId & "": CommandExecute
        End If
    End If
    
    DeleteTemporary
    ClearBox
    ClearFrame
    ClearItemBox
    BoxState False
    ButtonState False
    lvwPO.ListItems.Clear: cmdSearch.Enabled = True: cmdNew.Enabled = True: cmdNew.SetFocus
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
   BoxState False
   ButtonState False
   ClearBox
   ClearItemBox
   ClearFrame
   DeleteTemporary
   LoadSearchList
   EncodeMode = "S": txtPOSearch.Text = "": cboPOSearch.SetFocus:
   cmdSearch.Enabled = True: cmdNew.Enabled = True: txtPOArea.Visible = False: txtPOPrs.Visible = True
End Sub
Private Sub cmdPrint_Click()
    PORemark = txtPORemark.Text
    If EncodeMode = "A" Or EncodeMode = "E" Then
      strsql = " Insert Into PODetails Select * From PODetailsTemp": CommandExecute
    End If
    
    Load DataEnvironment1
    If DataEnvironment1.rsCommand7.State <> 0 Then DataEnvironment1.rsCommand7.Close
         ReportPO.Refresh
    If ReportPO.Visible = False Then ReportPO.Show
   
    If EncodeMode = "S" Then
        DeleteTemporary
    End If
        
    ClearBox
    ClearItemBox
    ClearFrame
    BoxState False
    ButtonState False
    cmdNew.Enabled = True: cmdSearch.Enabled = True

End Sub
Private Sub cmdAddClose_Click()
    frameItems.Visible = False: cmdAdd.SetFocus
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
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

Private Sub txtPODate_GotFocus()
   txtPODate.SelLength = Len(txtPODate.Text)
End Sub
Private Sub TxtPOdate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtPOArea.Visible = True: txtPOArea.SetFocus
   End If
End Sub
Private Sub txtPOArea_KeyPress(KeyAscii As Integer)
  If KeyAscii = 48 Then
     txtPOArea.Text = "DCO": txtPOPrs.Text = "DCO-PRS": txtPOArea_Click
  ElseIf KeyAscii = 49 Then
    txtPOArea.Text = "B1": txtPOPrs.Text = "B1-PRS": txtPOArea_Click
  ElseIf KeyAscii = 50 Then
    txtPOArea.Text = "B2": txtPOPrs.Text = "B2-PRS": txtPOArea_Click
  ElseIf KeyAscii = 51 Then
    txtPOArea.Text = "B3": txtPOPrs.Text = "B3-PRS": txtPOArea_Click
  ElseIf KeyAscii = 52 Then
    txtPOArea.Text = "B4": txtPOPrs.Text = "B4-PRS": txtPOArea_Click
  ElseIf KeyAscii = 53 Then
    txtPOArea.Text = "PFI": txtPOPrs.Text = "PFI-PRS": txtPOArea_Click
  End If
End Sub
Private Sub txtPOArea_Click()
  txtPOPrs.Text = txtPOArea.Text & "-PRS-"
  txtPOArea.Visible = False: txtPOPrs.Visible = True: txtPOPrs.SetFocus: txtPOPrs.SelStart = 9
End Sub
Private Sub txtPOPrs_DblClick()
 txtPOArea.Visible = True: txtPOPrs.Visible = False: txtPOArea.SetFocus
End Sub
Private Sub txtPOPrs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtPOPrs.Text = "" Then
       txtPOPrs.Text = "-"
    End If
      txtPOSupplier.SetFocus
End If
End Sub
Private Sub txtPOSupplier_GotFocus()
   txtPOSupplier.Text = "": txtPOSupplier.SelLength = Len(txtPOSupplier.Text)
End Sub
Private Sub txtPOSupplier_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii > 0 Then
      txtPOSupplier_DblClick
    End If
End Sub
Private Sub txtPOSupplier_DblClick()
   SetLvwPoSuppliers
   BoxState False
   LoadSuppliers
End Sub
Private Sub LoadSuppliers()
On Error GoTo LocalError
lvwPOSuppliers.ListItems.Clear
    strsql = "SELECT * from Suppliers ORDER BY SupName": CommandExecute
    With mmsADORst
        Do Until .EOF
            Set POLI = lvwPOSuppliers.ListItems.Add(, , !SupName & "")
                POLI.SubItems(1) = !SupId & ""
            .MoveNext
        Loop
     End With
LocalError: Exit Sub
End Sub
'------------   S U P P L I E R S   L I S T V I E W   -------------------
Private Function GetAdress() As String
     strsql = "SELECT SupAddress FROM Suppliers WHERE SupId = " & SupId & "": CommandExecute
     GetAdress = mmsADORst!SupAddress
End Function
Private Function GetAdress2() As String
     strsql = "SELECT SupAddress FROM Suppliers WHERE SupName like '" & POSup & "'": CommandExecute
     GetAdress2 = mmsADORst!SupAddress
End Function
Private Sub lvwPOSuppliers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      SupId = lvwPOSuppliers.SelectedItem.SubItems(1)
      txtPOSupplier.Text = lvwPOSuppliers.SelectedItem.Text
      txtPOSupAd.Text = GetAdress
    End With
End Sub
Private Sub lvwPOSuppliers_DblClick()
    frameSuppliers.Visible = False
    BoxState True
    txtPOTerms.SetFocus
End Sub
Private Sub lvwPOSuppliers_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddSupplier_Click
   End If
End Sub
Private Sub lvwPOSuppliers_KeyPress(KeyAscii As Integer)
    lvwPOSuppliers_DblClick
End Sub
Private Sub lvwPOSuppliers_GotFocus()
On Error GoTo LocalError
      SupId = lvwPOSuppliers.SelectedItem.SubItems(1)
      txtPOSupplier.Text = lvwPOSuppliers.SelectedItem.Text
      txtPOSupAd.Text = GetAdress
LocalError: Exit Sub
End Sub

'-----------------------------------------------------------------------------
Private Sub txtPOTerms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If txtPOTerms.Text = "" Then
          txtPOTerms.Text = "COD"
       End If
      txtPOWork.SetFocus
    End If
End Sub
Private Sub txtPOWOrk_GotFocus()
    strsql = "SELECT * FROM Transact ORDER BY TransactID": CommandExecute
    With mmsADORst
    txtPOWork.Clear
        Do While Not .EOF
            txtPOWork.AddItem ![TransactName]: .MoveNext
        Loop
    End With
End Sub
Private Sub txtPOWOrk_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       If txtPOWork.Text = "" Then
          txtPOWork.Text = "-"
       End If
       txtPOEquip.SetFocus
   End If
End Sub
Private Sub txtPOEquip_GotFocus()
    strsql = "SELECT * FROM Equipment ORDER BY EquipName": CommandExecute
    With mmsADORst
        txtPOEquip.Clear
        Do While Not .EOF
            txtPOEquip.AddItem ![EquipName]: .MoveNext
        Loop
    End With
    txtPOEquip.ListIndex = 0
End Sub
Private Sub txtPOEquip_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       txtPORemark.SetFocus
   End If
End Sub
Private Sub txtPORemark_GotFocus()
   If txtPOWork.Text = "REPAIR AND MAINTENANCE" Then
     txtPORemark.Text = "MAINTENANCE FOR " & txtPOEquip.Text
   Else
     txtPORemark.Text = txtPOWork.Text & " FOR "
   End If
   txtPORemark.SelStart = Len(txtPORemark.Text) + 1: txtPORemark.SelLength = Len(txtPORemark.Text)
End Sub
Private Sub txtPORemark_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
       If txtPORemark.Text = "" Then
          txtPORemark.Text = "-"
       End If
       LoadItemsList
    End If
   If Not IsDate(txtPODate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtPODate = Format$(Now, "mm/dd/yyyy"): txtPODate.SetFocus
  End If
End Sub
'------------------------------

'---------------------------------
Private Sub txtPOUnit_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
       txtPOCost.SetFocus
   End If
End Sub
Private Sub txtPOQty_GotFocus()
 txtPOQty.SelLength = Len(txtPOQty.Text)
End Sub
Private Sub txtPOQty_Change()
  TxtVal = txtPOQty.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPOQty.Text = ""
  End If
End Sub
Private Sub txtPOQty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtPOCost.SetFocus:  GetItemAmount
   End If
End Sub
Private Sub txtPOCost_GotFocus()
   If EncodeMode = "A" Then
    txtPOCost.Text = "0.00"
   End If
   txtPOCost.SelLength = Len(txtPOCost.Text)
End Sub
Private Sub txtPOCost_Change()
  TxtVal = txtPOCost.Text
  If IsNumeric(TxtVal) Then
    NumVal = TxtVal
  Else
    txtPOCost.Text = ""
  End If
End Sub
Private Sub txtPOCost_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not DataItemValidation Then
        Exit Sub
      End If
        cmdSavePODetails.SetFocus: GetItemAmount
   End If
End Sub
Private Sub GetItemAmount()
On Error GoTo LocalError
    txtPOCost.Text = Format$(txtPOCost, "#,###.#0"): txtPOQty.Text = Format$(txtPOQty, "#,###.#0"):
    ItemAmount = CDbl(txtPOCost.Text) * CDbl(txtPOQty.Text): txtPOAmount.Text = Format$(ItemAmount, "#,###.#0"):
LocalError: Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                      S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboPOSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtPOSearch.SetFocus
   End If
End Sub
Private Sub cboPOSearch_GotFocus()
  cboPOSearch.Text = "NUMBER"
End Sub
Private Sub txtPOSearch_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     lvwPOSearch.SetFocus
  End If
End Sub
Private Sub txtPOSearch_Change()
On Error GoTo LocalError
    If cboPOSearch.Text = "NUMBER" Then
       strsql = "Select * from PODetails where POID like '" & txtPOSearch.Text & "%'" & "Order by PONum"
    ElseIf cboPOSearch.Text = "SUPPLIER" Then
       strsql = "Select * from PODetails where POSupplier like '" & txtPOSearch.Text & "%'" & "Order by POSupplier"
    End If
    CommandExecute
    LoadPOSearch
LocalError: Exit Sub
End Sub
'---------- S E A R C H   L I S T V I E W-----
Private Sub lvwPOSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
       SearchLoad
    End With
End Sub
Private Sub lvwPOSearch_DblClick()
   frameSearch.Visible = False: cmdPrint.Enabled = True
   BoxState False
   SaveTemporary
   LoadPODetails
End Sub
Private Sub lvwPOSearch_KeyPress(KeyAscii As Integer)
On Error GoTo LocalError
   If KeyAscii = 13 Then
       lvwPOSearch_DblClick
       POResult = lvwPOSearch.SelectedItem.Text
   End If
LocalError: Exit Sub
End Sub
Private Sub lvwPOSearch_GotFocus()
On Error GoTo LocalError
    SearchLoad
LocalError: Exit Sub
End Sub
Private Sub SearchLoad()
      txtPONum.Text = lvwPOSearch.SelectedItem.Text: POId = txtPONum.Text
      txtPODate.Text = Format$(lvwPOSearch.SelectedItem.SubItems(1), "mm/dd/yyyy")
      txtPOPrs.Text = lvwPOSearch.SelectedItem.SubItems(2)
      POSup = lvwPOSearch.SelectedItem.SubItems(3): txtPOSupplier.Text = POSup: txtPOSupAd.Text = GetAdress2
      txtPORemark.Text = lvwPOSearch.SelectedItem.SubItems(4)
      txtPOWork2.Visible = True: txtPOWork2.Text = lvwPOSearch.SelectedItem.SubItems(5)
      'txtPOEquip.Text = lvwPOSearch.SelectedItem.SubItems(6)
      txtPOTotal.Text = lvwPOSearch.SelectedItem.SubItems(7)
End Sub
Private Sub LoadSearchList()
    SetLvwPoSearch
    strsql = "SELECT DISTINCT PONum, PODate, POPrs, POSupplier, PORemark, POWork, POEquip, PoTotal FROM PODetails ORDER BY PONum": CommandExecute
    LoadPOSearch
End Sub
Private Sub LoadPOSearch()
On Error GoTo LocalError
    lvwPOSearch.ListItems.Clear
    With mmsADORst
    Do Until .EOF
        Set POLI = lvwPOSearch.ListItems.Add(, , !PONum & "")
            POLI.SubItems(1) = !PODate & ""
            POLI.SubItems(2) = !POPrs & ""
            POLI.SubItems(3) = !POSupplier & ""
            POLI.SubItems(4) = !PORemark & ""
            POLI.SubItems(5) = !POWork & ""
            POLI.SubItems(6) = !POEquip & ""
            POLI.SubItems(7) = !POTotal & ""
            .MoveNext
    Loop
    End With
LocalError: Exit Sub
End Sub
'------------------------------------------------------------------------------------
'                  I T E M S   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtItemSearch_Change()
    strsql = "Select * from Materials where ItemName like '" & txtItemSearch.Text & "%' Order by ItemName": CommandExecute
    lvwPOItems.ListItems.Clear
    With mmsADORst
    Do Until .EOF
        Set POLI = lvwPOItems.ListItems.Add(, , !ItemGroup & "")
            POLI.SubItems(1) = !ItemCode & ""
            POLI.SubItems(2) = !ItemName & ""
            POLI.SubItems(4) = !Unit & ""
            POLI.SubItems(5) = CStr(!ItemId)
            .MoveNext
    Loop
    End With
End Sub
Private Sub txtItemSearch_Click()
    txtItemSearch.Text = ""
End Sub
Private Sub txtItemSearch_GotFocus()
   txtItemSearch.Text = "": txtItemSearch.SelLength = Len(txtItemSearch.Text)
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
    If KeyAscii = 13 Then
      lvwPOItems.SetFocus
    End If
End Sub
'---------- I T E M S  L I S T V I E W ----------
Private Sub LoadItemSearch()
      ItemGroup = lvwPOItems.SelectedItem.Text
      ItemCode = lvwPOItems.SelectedItem.SubItems(1)
      txtPOGroup.Text = ItemGroup & "-" & ItemCode
      txtPOItem.Text = lvwPOItems.SelectedItem.SubItems(2)
      txtPOUnit.Text = lvwPOItems.SelectedItem.SubItems(4)
End Sub
Private Sub lvwPOItems_GotFocus()
On Error GoTo LocalError
      LoadItemSearch
LocalError:
    Exit Sub
End Sub
Private Sub lvwPOItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      LoadItemSearch
    End With
End Sub
Private Sub lvwPOItems_KeyPress(KeyAscii As Integer)
   lvwPOItems_DblClick
End Sub
Private Sub lvwPOItems_DblClick()
    frameItems.Visible = False: frameItemAdd.Visible = True
    frameItemAdd.Top = lvwPO.Top - 50: frameItemAdd.Left = 7500: frameItemAdd.Height = 5000: frameItemAdd.Width = 10500
    BoxState False
    ItemBoxState True
    txtPOQty.SetFocus
End Sub
Private Sub LoadItemsList()
    If Not DataValidation Then
       Exit Sub
    End If
    SetLvwPoItems
    lvwPOItems.ListItems.Clear
    strsql = "SELECT * From Materials Order By ItemName": CommandExecute
    With mmsADORst
        Do Until .EOF
            Set POLI = lvwPOItems.ListItems.Add(, , !ItemGroup & "")
                POLI.SubItems(1) = !ItemCode & ""
                POLI.SubItems(2) = !ItemName & ""
                POLI.SubItems(4) = !Unit & ""
                POLI.SubItems(5) = CStr(!ItemId)
                .MoveNext
        Loop
    End With

LocalError: Exit Sub
End Sub

'------------------------------------------------------------------------------------
'                   S U P P L I E R   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub txtSupSearch_Change()
    strsql = "Select * from Suppliers where SupName like '" & txtSupSearch.Text & "%' Order by SupName": CommandExecute
    lvwPOSuppliers.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set POLI = lvwPOSuppliers.ListItems.Add(, , !SupName & "")
                POLI.SubItems(1) = !SupId
           .MoveNext
        Loop
     End With
End Sub
Private Sub txtSupSearch_Click()
    txtSupSearch.Text = ""
End Sub
Private Sub txtSupSearch_GotFocus()
   txtSupSearch.Text = ""
   txtSupSearch.SelLength = Len(txtSupSearch.Text)
End Sub
Private Sub txtSupSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwPOSuppliers.SetFocus
   End If
End Sub

'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub lvwPO_DblClick()
    If Not lvwPO.SelectedItem Is Nothing Then
        If EncodeMode = "A" Then
            EncodeMode = "E"
            ClearItemBox
            frameItemAdd.Visible = True
            POItemNo = lvwPO.SelectedItem
            txtPOGroup.Text = lvwPO.SelectedItem.SubItems(1)
            txtPOItem.Text = lvwPO.SelectedItem.SubItems(2)
            txtPOQty.Text = lvwPO.SelectedItem.SubItems(3)
            txtPOUnit.Text = lvwPO.SelectedItem.SubItems(4)
            txtPOCost.Text = lvwPO.SelectedItem.SubItems(5)
            txtPOAmount.Text = lvwPO.SelectedItem.SubItems(6)
            txtPOQty.SetFocus
        End If
   End If
End Sub
Private Sub DeleteEntry()
    ItemNoDel = lvwPO.SelectedItem
    ItemCodeDel = lvwPO.SelectedItem.SubItems(1)
    ItemQtyDel = lvwPO.SelectedItem.SubItems(3)
    
    strsql = "Delete From MRSDetailsTemp where MRSItemNo like " & ItemNoDel & "": CommandExecute
    strsql = "Update MRSDetailsTemp SET MRSItemNo = MRSItemNo - 1 WHERE MRSItemNo > " & ItemNoDel & "": CommandExecute

    POItemNo = POItemNo - 1
    LoadPODetails
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
Private Sub LoadPODetails()
On Error GoTo LocalError

    GetPOTotal
    txtPOTotal.Text = Format$(POTotal, "#,###.#0"): NumWords: InsertPOTotal
      
    strsql = "SELECT * from PODetailsTemp ORDER BY POItemNo ": CommandExecute
    lvwPO.ListItems.Clear
      With mmsADORst
        Do Until .EOF
            Set POLI = lvwPO.ListItems.Add(, , !POItemNo & "")
            POLI.SubItems(1) = !POCode & ""
            POLI.SubItems(2) = !POItem & ""
            POLI.SubItems(3) = !POQty & ""
            POLI.SubItems(4) = !POUnit & ""
            POLI.SubItems(5) = !POCost & ""
            POLI.SubItems(6) = !POAmount & ""
            .MoveNext
        Loop
      End With
    cmdAdd.SetFocus
    
LocalError: Exit Sub
End Sub
Private Sub GetPOTotal()
    strsql = " SELECT SUM(POAmount) as SubTotal FROM PODetailsTemp ": CommandExecute
    POTotal = mmsADORst.Fields!Subtotal: txtPOTotal.Text = Format$(POTotal, "#,###.#0")
End Sub
Private Sub InsertPOTotal()
On Error GoTo LocalError
    strsql = "Update PODetailsTemp SET POTotal = '" & txtPOTotal.Text & "' where PONum like '" & txtPONum.Text & "'": CommandExecute
LocalError:
    Exit Sub
End Sub
Private Sub SaveTemporary()
On Error GoTo LocalError

  If txtPOArea.Text = "DCO" Or txtPOWork.Text = "REPAIR AND MAINTENANCE" Then
     txtPOStatus.Text = "SERVED"
  End If

  If EncodeMode = "A" Then
    strsql = POSaveTemp: CommandExecute
  End If
  
  If EncodeMode = "S" Then
     strsql = " Insert Into PODetailsTemp Select * From PODetails Where PONum like '" & txtPONum.Text & "' Order By POItemNo ": CommandExecute
     cmdCancel.Enabled = True
  End If
  
LocalError: Exit Sub
End Sub
Private Function GetNextPOID() As Long
   strsql = "SELECT MAX(POID) AS MaxID FROM PODetails": CommandExecute
       If mmsADORst.EOF Then
           GetNextPOID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextPOID = 1
       Else
           GetNextPOID = mmsADORst!MaxID + 1
       End If
End Function
Private Sub DeleteTemporary()
On Error GoTo LocalError
    strsql = "Delete From PODetailsTemp": CommandExecute
LocalError: Exit Sub
End Sub
Public Function DataValidation() As Boolean
DataValidation = False
    If txtPODate.Text = "" Then
        MsgBox "Fill-up PO Date.", vbExclamation, "PO Date Required"
        txtPODate.SetFocus: Exit Function
    End If
    If txtPOPrs.Text = "" Then
        MsgBox "Fill-up PRS Number.", vbExclamation, "PO Required"
        txtPOArea.SetFocus: Exit Function
    End If
    If txtPOSupplier.Text = "" Then
        MsgBox "Fill-up Supplier Name.", vbExclamation, "Supplier Required"
        txtPOSupplier.SetFocus: Exit Function
    End If
    If txtPOWork.Text = "" Then
        MsgBox "Fill-up transaction.", vbExclamation, "Supplier Required"
        txtPOWork.SetFocus: Exit Function
    End If
    If txtPOEquip.Text = "" Then
       txtPOEquip.Text = "-"
    End If
DataValidation = True
End Function
Public Function DataItemValidation() As Boolean
On Error GoTo LocalError
  DataItemValidation = False
    If txtPOItem.Text = "" Or txtPOUnit.Text = "" Then
        MsgBox "No Item Name.", vbExclamation, "Item Required"
        txtPOItem.SetFocus: Exit Function
    End If
    If txtPOQty.Text = "" Or CDbl(txtPOQty.Text) = 0 Or txtPOQty = Null Then
        MsgBox "No Quantity of Item.", vbExclamation, "Quantity Required"
        txtPOQty.SetFocus: Exit Function
    End If
  DataItemValidation = True
LocalError: Exit Function
End Function

'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HC0C0C0
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0C0C0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdSavePODetails_GotFocus()
   cmdSavePODetails.BackColor = &HC0C0C0
End Sub
Private Sub cmdSavePODetails_LostFocus()
   cmdSavePODetails.BackColor = &H8000000F
End Sub
Private Sub cmdEdit_GotFocus()
   cmdEdit.BackColor = &HC0C0C0
End Sub
Private Sub cmdEdit_LostFocus()
   cmdEdit.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0C0C0
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub
Private Sub cmdSave_GotFocus()
   'cmdSave.BackColor = &H00C0C0C0&
End Sub
Private Sub cmdSave_LostFocus()
   'cmdSave.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0C0C0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HC0C0C0
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


