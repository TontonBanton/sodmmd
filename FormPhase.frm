VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormArea 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PHASE LIBRARY"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10663.67
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAreaCode 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      MaxLength       =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1097
      Left            =   0
      TabIndex        =   2
      Top             =   9623
      Width           =   8745
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   2000
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   2200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   2000
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&ADD"
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
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   2000
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
         Left            =   6400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   2000
      End
   End
   Begin VB.TextBox txtAddPhase 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   209
      MaxLength       =   50
      TabIndex        =   1
      Top             =   8962
      Width           =   8212
   End
   Begin MSComctlLib.ListView lvwPhase 
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   15584
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
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
Attribute VB_Name = "FormArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private mstrSQL                As String
Private EncodeMode             As String
Private ButtonPress            As String
Private Search                 As Boolean
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwPhase
   LoadPhase
   'lblComp.Caption = FormMainMenu.lblComp.Caption: lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub
Private Sub txtAddPhase_GotFocus()
   txtAddPhase.SelStart = 0
   txtAddPhase.SelLength = Len(txtAddPhase.Text)
End Sub

Private Sub txtAddPhase_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      cmdCancel_Click
      Exit Sub
   ElseIf KeyAscii = 13 Then
      cmdSave.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.SetFocus
   End If
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ButtonState False
    BoxState True
    ClearBox
    txtAddPhase.SetFocus
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
   If Not DataValidation Then
      Exit Sub
   End If
      If EncodeMode = "A" Then

         lngIDField = GetNextAreaID
         
         strsql = "INSERT INTO Area (  AreaID"
         strsql = strsql & "            , AreaName"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtAddPhase.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwPhase.ListItems.Add(, , txtAddPhase.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(1) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwPhase.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwPhase.SelectedItem.SubItems(1))

         strsql = "UPDATE Area SET "
         strsql = strsql & "  AreaName        = '" & Replace$(txtAddPhase.Text, "'", "''") & "'"
         strsql = strsql & " WHERE AreaID = " & lngIDField
 
         lvwPhase.SelectedItem.Text = txtAddPhase.Text
         PopulateItem lvwPhase.SelectedItem
      End If
 
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwPhase.Enabled = True
    lvwPhase.SetFocus
    ButtonState True
    
    'txtItemSearch.Text = ""
    
End Sub
Private Sub cmdCancel_Click()
Clear
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     Unload Me
     FormMainMenu.Show
     FormMainMenu.Enabled = True
 Else
     Exit Sub
 End If

End Sub
Private Sub Clear()
 txtAddPhase.Text = ""
 cmdAdd.Enabled = True
 cmdExit.Enabled = True
 lvwPhase.Enabled = True
End Sub
'------------ F O C U S ---------------
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdSave_GotFocus()
   cmdSave.BackColor = &HC0FFC0
End Sub
Private Sub cmdSave_LostFocus()
   cmdSave.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub

'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub SetlvwPhase()
    With lvwPhase
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", .Width * 0.93
        .ColumnHeaders.Add , , "", Width * 0#
    End With
End Sub
Private Sub lvwPhase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'LoadMaterials
    With lvwPhase
        If (.Sorted) And (ColumnHeader.SubItemIndex = .SortKey) Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .Sorted = True
            .SortKey = ColumnHeader.SubItemIndex
            .SortOrder = lvwAscending
        End If
        .Refresh
    End With
        
    If Not lvwPhase.SelectedItem Is Nothing Then
        lvwPhase.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwPhase_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtAddPhase.Text = .Text
        TxtAreaCode.Text = .SubItems(1)
     End With
End Sub
Private Sub lvwPhase_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub
Private Sub lvwPhase_DblClick()
    ButtonState False
    BoxState True
    'txtProduct.Enabled = False
    'txtClassification.SetFocus
    If EncodeMode = "S" Then
           With lvwPhase
           If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwPhase_ItemClick .SelectedItem
                  End If
           End With
    Else
            EncodeMode = "U"
    End If
End Sub
'--------------------------------------------
'                   F U N C T I O N S
'--------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        mstrSQL = "select * from Suppliers"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open mstrSQL, mmsADOConn, adOpenDynamic, adLockOptimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub LoadPhase()
    Dim strsql        As String
    Dim PhaseLI  As ListItem
                                 
    strsql = "SELECT * FROM Area ORDER BY AreaName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwPhase.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set PhaseLI = lvwPhase.ListItems.Add(, , !AreaName & "")
            PhaseLI.SubItems(1) = !AreaID & ""
            .MoveNext
        Loop
     End With
    
              With lvwPhase
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     'lvwPhase_ItemClick .SelectedItem
                  End If
              End With
     
    Set PhaseLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = TxtAreaCode.Text
    End With
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtAddPhase.Text = "" Then
         MsgBox "FILL UP PHASE", vbExclamation, "PHASE"
         txtAddPhase.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function GetNextAreaID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(AreaID) AS MaxID FROM Area"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextAreaID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextAreaID = 1
       Else
           GetNextAreaID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub ClearBox()
    txtAddPhase.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtAddPhase.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwPhase.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub

