VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormEquip 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQUIPMENT LIBRARY"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10572.79
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAddEquip 
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
      Height          =   500
      Left            =   200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   8880
      Width           =   8212
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1097
      Left            =   0
      TabIndex        =   3
      Top             =   9503
      Width           =   8745
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
         TabIndex        =   7
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
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   150
         Width           =   2000
      End
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
         TabIndex        =   6
         Top             =   150
         Width           =   2000
      End
   End
   Begin VB.TextBox TxtEquipCode 
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwEquip 
      Height          =   8715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   15372
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
Attribute VB_Name = "FormEquip"
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
   SetlvwEquip
   LoadEquip
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub
Private Sub txtAddEquip_GotFocus()
   txtAddEquip.SelStart = 0
   txtAddEquip.SelLength = Len(txtAddEquip.Text)
End Sub

Private Sub txtAddEquip_KeyPress(KeyAscii As Integer)
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
    txtAddEquip.SetFocus
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
   If Not DataValidation Then
      Exit Sub
   End If
      If EncodeMode = "A" Then

         lngIDField = GetNextEquipID
         
         strsql = "INSERT INTO Equipment (  EquipID"
         strsql = strsql & "            , EquipName"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtAddEquip.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwEquip.ListItems.Add(, , txtAddEquip.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(1) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwEquip.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwEquip.SelectedItem.SubItems(1))

         strsql = "UPDATE Equip SET "
         strsql = strsql & "  EquipName        = '" & Replace$(txtAddEquip.Text, "'", "''") & "'"
         strsql = strsql & " WHERE EquipID = " & lngIDField
 
         lvwEquip.SelectedItem.Text = txtAddEquip.Text
         PopulateItem lvwEquip.SelectedItem
      End If
 
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwEquip.Enabled = True
    lvwEquip.SetFocus
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
 txtAddEquip.Text = ""
 cmdAdd.Enabled = True
 cmdExit.Enabled = True
 lvwEquip.Enabled = True
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
Private Sub SetlvwEquip()
    With lvwEquip
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", .Width * 0.93
        .ColumnHeaders.Add , , "", Width * 0#
    End With
End Sub
Private Sub lvwEquip_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'LoadMaterials
    With lvwEquip
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
        
    If Not lvwEquip.SelectedItem Is Nothing Then
        lvwEquip.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwEquip_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtAddEquip.Text = .Text
        TxtEquipCode.Text = .SubItems(1)
     End With
End Sub
Private Sub lvwEquip_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub
Private Sub lvwEquip_DblClick()
    ButtonState False
    BoxState True
    'txtProduct.Enabled = False
    'txtClassification.SetFocus
    If EncodeMode = "S" Then
           With lvwEquip
           If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwEquip_ItemClick .SelectedItem
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
Private Sub LoadEquip()
    Dim strsql        As String
    Dim EquipLI  As ListItem
                                 
    strsql = "SELECT * FROM Equipment ORDER BY EquipName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwEquip.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set EquipLI = lvwEquip.ListItems.Add(, , !EquipName & "")
            EquipLI.SubItems(1) = !EquipID & ""
            .MoveNext
        Loop
     End With
    
              With lvwEquip
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     'lvwEquip_ItemClick .SelectedItem
                  End If
              End With
     
    Set EquipLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = TxtEquipCode.Text
    End With
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtAddEquip.Text = "" Then
         MsgBox "FILL UP Equipment", vbExclamation, "Equipment"
         txtAddEquip.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function GetNextEquipID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(EquipID) AS MaxID FROM Equipment"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextEquipID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextEquipID = 1
       Else
           GetNextEquipID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub ClearBox()
    txtAddEquip.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtAddEquip.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwEquip.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub



