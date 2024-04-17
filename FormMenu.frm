VERSION 5.00
Begin VB.Form FormMainMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  MATERIALS MANAGEMENT SYSTEM  "
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   ForeColor       =   &H80000008&
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameForms 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2630
      Begin VB.CommandButton cmdPO 
         Caption         =   "PURCHASE ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   2450
      End
      Begin VB.CommandButton cmdMIS 
         Caption         =   "ISSUANCE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2910
         Width           =   2450
      End
      Begin VB.CommandButton cmdPRS 
         Caption         =   "REQUEST"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1500
         Width           =   2450
      End
      Begin VB.CommandButton cmdMRR 
         Caption         =   "RECEIVING"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2205
         Width           =   2450
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "TRANSMITTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   795
         Width           =   2450
      End
   End
   Begin VB.Frame frameLib 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   2760
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   2500
      Begin VB.CommandButton cmdItems 
         Caption         =   "ITEMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   2300
      End
      Begin VB.CommandButton cmdSuppliers 
         Caption         =   "SUPPLIERS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   795
         Width           =   2300
      End
      Begin VB.CommandButton cmdTransact 
         Caption         =   "TRANSACTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1500
         Width           =   2300
      End
      Begin VB.CommandButton cmdPhase 
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2205
         Width           =   2300
      End
      Begin VB.CommandButton cmdEquip 
         Caption         =   "EQUIPMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2910
         Width           =   2300
      End
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "ADJUSTMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame frameTools 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2350
      Left            =   5280
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   2500
      Begin VB.CommandButton cmdConvert 
         Caption         =   "CONVERSION TOOLS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   2300
      End
      Begin VB.CommandButton cmdSettings 
         Caption         =   "SETTINGS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1500
         Width           =   2300
      End
      Begin VB.CommandButton cmdCalculator 
         Caption         =   "CALCULATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   800
         Width           =   2300
      End
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "&TOOLS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14265
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdLib 
      Caption         =   "&LIBRARIES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdForms 
      Caption         =   "&FORMS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   2500
   End
   Begin VB.Timer Timer1 
      Left            =   12000
      Top             =   360
   End
   Begin VB.Label lblArea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA"
      Height          =   195
      Left            =   1080
      TabIndex        =   15
      Top             =   515
      Visible         =   0   'False
      Width           =   435
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
      Left            =   1080
      TabIndex        =   14
      Top             =   68
      Width           =   1125
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Location - Adress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   13
      Top             =   272
      Width           =   1560
   End
   Begin VB.Image Image5 
      Height          =   5100
      Left            =   4150
      Picture         =   "FormMenu.frx":1E72
      Stretch         =   -1  'True
      Top             =   3100
      Width           =   8650
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   16935
   End
   Begin VB.Label lblClock 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   16080
      TabIndex        =   9
      Top             =   360
      Width           =   570
   End
   Begin VB.Label lblToday 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   16080
      TabIndex        =   8
      Top             =   120
      Width           =   570
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   0
      Picture         =   "FormMenu.frx":3AE93
      Stretch         =   -1  'True
      Top             =   29
      Width           =   990
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   825
      Left            =   -120
      TabIndex        =   6
      Top             =   29
      Width           =   17040
   End
   Begin VB.Image Image4 
      Height          =   6765
      Left            =   3884
      Picture         =   "FormMenu.frx":4466F
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   9135
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Transaction &Forms"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuPO 
         Caption         =   "Purchase Order"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Transmittal"
      End
      Begin VB.Menu mnuPRS 
         Caption         =   "Purchase Requisition"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuMRR 
         Caption         =   "Materials Receiving"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMIS 
         Caption         =   "Materials Issuance"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Libraries"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTransact 
         Caption         =   "Transact"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuArea 
         Caption         =   "Area"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEquip 
         Caption         =   "Equipments"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuConvet 
         Caption         =   "Conversion"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FormMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String

Private Sub Form_Activate()
On Error GoTo LocalError
  ConnectToDB
  mmsAdoCmd.CommandText = "Select * From Settings"
  Set mmsADORst = mmsAdoCmd.Execute
  lblComp.Caption = mmsADORst.Fields("Company")
  lblHeader.Caption = mmsADORst.Fields("Location")
  lblArea.Caption = mmsADORst.Fields("Area")
LocalError: Exit Sub
End Sub
Private Sub Form_Load()
  Load Me
  Timer1.Interval = 100
  FrameForms.Visible = False
  frameLib.Visible = False
  ConnectToDB
  Form_Activate
End Sub
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
Private Sub Timer1_Timer()
   lblToday.Caption = Format$(Now, "ddd, mmm/dd/yyyy")
   lblClock.Caption = Format$(Now, "hh:mm AM/PM")
End Sub
'------------------------------------------------------------------------
'                      B U T T O N S   E V E N T S
'-------------------------------------------------------------------------
Private Sub cmdForms_Click()
   FrameForms.Visible = True: cmdPO.SetFocus
   frameLib.Visible = False: frameTools.Visible = False
End Sub
Private Sub cmdLib_Click()
  frameLib.Visible = True: cmdItems.SetFocus
  FrameForms.Visible = False: frameTools.Visible = False
End Sub
Private Sub CmdTools_Click()
  frameTools.Visible = True: cmdConvert.SetFocus
  frameLib.Visible = False: FrameForms.Visible = False
End Sub
Private Sub CmdReport_Click()
  frameLib.Visible = False: FrameForms.Visible = False
  FormReport.Show
End Sub
Private Sub CmdExit_Click()
  If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
       End
  Else
       Exit Sub
  End If
End Sub
Private Sub cmdPO_Click()
   FormPO.Show
End Sub
Private Sub cmdTrans_Click()
  FormTransfer.Show
End Sub
Private Sub cmdMRR_Click()
  FormMRR.Show
End Sub
Private Sub cmdMRR_KeyPress(KeyAscii As Integer)
  cmdLib_GotFocus
End Sub
Private Sub cmdMIS_Click()
   FormMIS.Show
End Sub
Private Sub cmdMIS_KeyPress(KeyAscii As Integer)
  cmdLib_GotFocus
End Sub
Private Sub cmdPRS_Click()
  FormPRS.Show
End Sub
Private Sub cmdPRS_KeyPress(KeyAscii As Integer)
  cmdLib_GotFocus
End Sub
Private Sub cmdEquip_Click()
  FormEquip.Show
End Sub
Private Sub cmdEquip_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdItems_Click()
   FormItems.Show
End Sub
Private Sub cmdItems_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdSuppliers_Click()
  FormSuppliers.Show
End Sub
Private Sub cmdSuppliers_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdTransact_Click()
   FormTransact.Show
End Sub
Private Sub cmdTransact_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdPhase_Click()
  FormArea.Show
End Sub
Private Sub cmdPhase_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdConvert_Click()
  MsgBox "Under Renovation", vbInformation, "Convert"
  Shell "explorer.exe http://www.unitconverters.net/"
End Sub
Private Sub cmdConvert_KeyPress(KeyAscii As Integer)
  cmdReport_GotFocus
End Sub
Private Sub cmdCalculator_Click()
 Shell "C:\Windows\System32\calc.exe"
End Sub
Private Sub cmdCalculator_KeyPress(KeyAscii As Integer)
  cmdReport_GotFocus
End Sub
Private Sub cmdSettings_Click()
   FormLog.txtEncodeMode.Text = "SETTINGS"
   FormLog.Show
End Sub
Private Sub cmdSettings_KeyPress(KeyAscii As Integer)
  cmdReport_GotFocus
End Sub
'------------ F O C U S ---------------
Private Sub cmdForms_GotFocus()
   cmdForms.BackColor = &HFFFFC0
   FrameForms.Visible = True
   frameLib.Visible = False
   frameTools.Visible = False
End Sub
Private Sub cmdForms_LostFocus()
   cmdForms.BackColor = &H8000000F
End Sub
Private Sub cmdLib_GotFocus()
   cmdLib.BackColor = &HFFFFC0
   frameLib.Visible = True
   FrameForms.Visible = False
   frameTools.Visible = False
End Sub
Private Sub cmdLib_LostFocus()
   cmdLib.BackColor = &H8000000F
End Sub
Private Sub cmdTools_GotFocus()
   cmdTools.BackColor = &HFFFFC0
   frameTools.Visible = True
   frameLib.Visible = False
   FrameForms.Visible = False
End Sub
Private Sub cmdTools_LostFocus()
   cmdTools.BackColor = &H8000000F
End Sub
Private Sub cmdReport_GotFocus()
   cmdReport.BackColor = &HFFFFC0
   frameLib.Visible = False
   FrameForms.Visible = False
   frameTools.Visible = False
End Sub
Private Sub cmdReport_LostFocus()
   cmdReport.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HFFFFC0
   frameLib.Visible = False
   FrameForms.Visible = False
   frameTools.Visible = False
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub cmdPO_GotFocus()
   cmdPO.BackColor = &HFFFFC0
End Sub
Private Sub cmdPO_LostFocus()
   cmdPO.BackColor = &H8000000F
End Sub
Private Sub cmdTrans_GotFocus()
   cmdTrans.BackColor = &HFFFFC0
End Sub
Private Sub cmdTrans_LostFocus()
   cmdTrans.BackColor = &H8000000F
End Sub
Private Sub cmdMRR_GotFocus()
   cmdMRR.BackColor = &HFFFFC0
End Sub
Private Sub cmdMRR_LostFocus()
   cmdMRR.BackColor = &H8000000F
End Sub
Private Sub cmdMIS_GotFocus()
   cmdMIS.BackColor = &HFFFFC0
End Sub
Private Sub cmdMIS_LostFocus()
   cmdMIS.BackColor = &H8000000F
End Sub
Private Sub cmdPRS_GotFocus()
   cmdPRS.BackColor = &HFFFFC0
End Sub
Private Sub cmdPRS_LostFocus()
   cmdPRS.BackColor = &H8000000F
End Sub
Private Sub cmdItems_GotFocus()
   cmdItems.BackColor = &HFFFFC0
End Sub
Private Sub cmdItems_LostFocus()
   cmdItems.BackColor = &H8000000F
End Sub
Private Sub cmdSuppliers_GotFocus()
   cmdSuppliers.BackColor = &HFFFFC0
End Sub
Private Sub cmdSuppliers_LostFocus()
   cmdSuppliers.BackColor = &H8000000F
End Sub
Private Sub cmdTransact_GotFocus()
   cmdTransact.BackColor = &HFFFFC0
End Sub
Private Sub cmdTransact_LostFocus()
   cmdTransact.BackColor = &H8000000F
End Sub
Private Sub cmdPhase_GotFocus()
   cmdPhase.BackColor = &HFFFFC0
End Sub
Private Sub cmdPhase_LostFocus()
   cmdPhase.BackColor = &H8000000F
End Sub
Private Sub cmdEquip_GotFocus()
   cmdEquip.BackColor = &HFFFFC0
End Sub
Private Sub cmdEquip_LostFocus()
   cmdEquip.BackColor = &H8000000F
End Sub
Private Sub cmdAdjust_GotFocus()
   'cmdAdjust.BackColor = &H00FFFFC0&
End Sub
Private Sub cmdAdjust_LostFocus()
   'cmdAdjust.BackColor = &H8000000F
End Sub
Private Sub cmdConvert_GotFocus()
   cmdConvert.BackColor = &HFFFFC0
End Sub
Private Sub cmdConvert_LostFocus()
   cmdConvert.BackColor = &H8000000F
End Sub
Private Sub cmdCalculator_GotFocus()
   cmdCalculator.BackColor = &HFFFFC0
End Sub
Private Sub cmdCalculator_LostFocus()
   cmdCalculator.BackColor = &H8000000F
End Sub
Private Sub cmdSettings_GotFocus()
   cmdSettings.BackColor = &HFFFFC0
End Sub
Private Sub cmdSettings_LostFocus()
   cmdSettings.BackColor = &H8000000F
End Sub

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


