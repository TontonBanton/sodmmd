VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormItems 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    ITEMS LIBRARY "
   ClientHeight    =   10500
   ClientLeft      =   2640
   ClientTop       =   375
   ClientWidth     =   12990
   Icon            =   "FormMat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtMCode 
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
      Left            =   9240
      MaxLength       =   15
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboUnit 
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
      ItemData        =   "FormMat.frx":08CA
      Left            =   11160
      List            =   "FormMat.frx":08D7
      TabIndex        =   3
      Top             =   8520
      Width           =   1740
   End
   Begin VB.TextBox txtDes 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   8520
      Width           =   8655
   End
   Begin VB.ComboBox CboGroup 
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
      ItemData        =   "FormMat.frx":08EA
      Left            =   120
      List            =   "FormMat.frx":0900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8520
      Width           =   2220
   End
   Begin VB.TextBox txtItemSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   4
      Top             =   9000
      Width           =   10500
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   0
      TabIndex        =   5
      Top             =   9600
      Width           =   15250
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
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
         Left            =   4350
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   2100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&SAVE"
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
         Left            =   2230
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   2100
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
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   150
         Width           =   2100
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   2100
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&DELETE"
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
         Left            =   6460
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   2100
      End
   End
   Begin MSComctlLib.ListView lvwMaterials 
      Height          =   8355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   14737
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label REORDER 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1440
      TabIndex        =   12
      Top             =   9100
      Width           =   810
   End
End
Attribute VB_Name = "FormItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private mstrSQL                As String
Private Search                 As Boolean

Private EncodeMode, ButtonPress, ItemCode       As String
Private ItemAmount                              As Double
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwMaterials
   LoadMaterials
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub

'---------------------------------------------------------------------------------
'                                   C O N T R O L S   E V E N T S
'---------------------------------------------------------------------------------
Private Sub txtDes_GotFocus()
   txtDes.SelStart = 0
   txtDes.SelLength = Len(txtDes.Text)
End Sub
Private Sub TxtDes_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
        cboUnit.SetFocus
  End If
End Sub
Private Sub CboGroup_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        txtDes.SetFocus
      End If
End Sub
Private Sub CboGroup_LostFocus()
Dim ItemGroup As String
Dim ItemId    As String

If CboGroup.Text = "" Then
   Exit Sub
Else
  If EncodeMode = "A" Then
     ItemGroup = CboGroup.Text
     ItemId = Format$(GetNextItemID, "000000")
     TxtMCode.Text = Mid$(ItemGroup, 1, 1) & ItemId
  End If
End If
End Sub
Private Sub CboUnit_GotFocus()
'If CboGroup.Text = "POL" Or CboGroup.Text = "CHEMICALS" Then
'    cboUnit.Text = "LTS"
'ElseIf CboGroup.Text = "MATERIALS" Then
'    cboUnit.Text = "PCS"
'ElseIf CboGroup.Text = "FERTILIZERS" Then
'    cboUnit.Text = "KLO"
'End If
End Sub
Private Sub CboUnit_Keypress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
     cmdSave.SetFocus
   ElseIf IsNumeric(Chr(KeyAscii)) Then
      SendKeys_
   End If
End Sub
'------------------------------------------------------------------------------------
'                            S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboItemSearch_GotFocus()
   'cboItemSearch.Text = "Name"
   'cboItemSearch.SelStart = 0
   'cboItemSearch.SelLength = Len(cboItemSearch.Text)
End Sub
Private Sub cboItemSearch_Click()
  'ButtonState False
  'cmdSave.Enabled = False
  lvwMaterials.Enabled = True
  Search = True
End Sub
Private Sub cboItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii < 255 Then
      SendKeys_
   End If
   If KeyAscii = 13 Then
      txtItemSearch.SetFocus
   End If
End Sub
Private Sub txtItemSearch_Change()
Dim strsql       As String
Dim MaterialsLI  As ListItem
On Error GoTo LocalError
    Search = True
    
       strsql = "Select * from Materials where ItemName like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemCode"
      
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVMaterials
LocalError:
    Exit Sub
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
    lvwMaterials.SetFocus
   End If
End Sub

'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ButtonState False
    BoxState True
    ClearBox
    CboGroup.SetFocus
End Sub
Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
End Sub
Private Sub cmdDelete_Click()

    If lvwMaterials.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        Exit Sub
    End If

    If MsgBox("Are you sure that you want to delete the item on the list " _
              , vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    ConnectToDB
    mmsAdoCmd.CommandText = "DELETE FROM Materials WHERE ItemCode =  '" & TxtMCode.Text & "'"
    mmsAdoCmd.Execute
    ClearBox
    LoadMaterials
End Sub
Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
 
On Error GoTo LocalError



   If Not DataValidation Then
      Exit Sub
   End If
   
        If EncodeMode = "A" Then
         lngIDField = GetNextItemID
    
         strsql = "INSERT INTO Materials (  ItemID"
         strsql = strsql & "            , ItemGroup"
         strsql = strsql & "            , ItemCode"
         strsql = strsql & "            , ItemName"
         strsql = strsql & "            , Unit"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(CboGroup.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(TxtMCode.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtDes.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwMaterials.ListItems.Add(, , txtDes.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(4) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwMaterials.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwMaterials.SelectedItem.SubItems(4))

         strsql = "UPDATE Materials SET "
         strsql = strsql & "  ItemGroup           = '" & Replace$(CboGroup.Text, "'", "''") & "'"
         strsql = strsql & ", Itemcode            = '" & Replace$(TxtMCode.Text, "'", "''") & "'"
         strsql = strsql & ", ItemName            = '" & Replace$(txtDes.Text, "'", "''") & "'"
         strsql = strsql & ", Unit                = '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & " WHERE ItemCode = '" & TxtMCode.Text & "'"

         lvwMaterials.SelectedItem.Text = txtDes.Text
         PopulateItem lvwMaterials.SelectedItem
    End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwMaterials.Enabled = True
    lvwMaterials.SetFocus
    ButtonState True
    

    txtItemSearch.Text = ""

LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwMaterials.Enabled = True
   txtItemSearch.Text = ""
    
    If lvwMaterials.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     Unload Me
 Else
     Exit Sub
 End If
End Sub
Private Sub cmdExit_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
End Sub
'------------ F O C U S ---------------
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdUpdate_GotFocus()
   'cmdUpdate.BackColor = &HC0FFC0
End Sub
Private Sub cmdUpdate_LostFocus()
   'cmdUpdate.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0FFC0
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
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
Private Sub cmdConvert_GotFocus()
   'cmdConvert.BackColor = &HC0FFC0
End Sub
Private Sub cmdConvert_LostFocus()
   'cmdConvert.BackColor = &H8000000F
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
Private Sub lvwMaterials_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'LoadMaterials
    With lvwMaterials
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
        
    If Not lvwMaterials.SelectedItem Is Nothing Then
        lvwMaterials.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwMaterials_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtDes.Text = .Text
        TxtMCode.Text = .SubItems(1)
        CboGroup.Text = .SubItems(2)
        cboUnit.Text = .SubItems(3)
     End With
End Sub
Private Sub lvwMaterials_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub
Private Sub lvwMaterials_DblClick()
    ButtonState False
    BoxState True
    'txtProduct.Enabled = False
    'txtClassification.SetFocus
    If EncodeMode = "S" Then
           With lvwMaterials
           If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwMaterials_ItemClick .SelectedItem
                  End If
           End With
    Else
            EncodeMode = "U"
    End If
End Sub

'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        mstrSQL = "select * from Materials"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open mstrSQL, mmsADOConn, adOpenDynamic, adLockOptimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub SetlvwMaterials()
    With lvwMaterials
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.5
        .ColumnHeaders.Add , , "Code", .Width * 0.15
        .ColumnHeaders.Add , , "Group", .Width * 0.15
        .ColumnHeaders.Add , , "Unit", .Width * 0.15
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = TxtMCode.Text
        .SubItems(2) = CboGroup.Text
        .SubItems(4) = cboUnit.Text
    End With
End Sub
Private Sub LoadMaterials()
    Dim strsql       As String
                          
    strsql = "SELECT * From Materials Order By ItemName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVMaterials
              With lvwMaterials
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwMaterials_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVMaterials()
Dim MaterialsLI  As ListItem
lvwMaterials.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwMaterials.ListItems.Add(, , !ItemName & "")
            MaterialsLI.SubItems(1) = !ItemCode & ""
            MaterialsLI.SubItems(2) = !ItemGroup & ""
            MaterialsLI.SubItems(3) = !Unit & ""
            MaterialsLI.SubItems(4) = CStr(!ItemId)
            .MoveNext
        Loop
     End With
    Set MaterialsLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub ClearBox()
    'CboGroup.Text = ""
    TxtMCode.Text = ""
    txtDes.Text = ""
    cboUnit.Text = ""
    txtItemSearch.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    CboGroup.Enabled = boxEnabled
    txtDes.Enabled = boxEnabled
    cboUnit.Enabled = boxEnabled
    txtItemSearch.Enabled = Not boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwMaterials.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdDelete.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub
Private Sub SendKeys_()
On Error GoTo LocalError
    'SendKeys "{left}"
    'SendKeys "{del}"
LocalError:
    Exit Sub
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtDes.Text = "" Then
        MsgBox "Fill-up Item Item Name", vbExclamation, "Item Required"
        CboGroup.SetFocus
        Exit Function
    End If
    If CboGroup.Text = "" Then
        MsgBox "Fill-up Item Group.", vbExclamation, "Item Group Required"
        CboGroup.SetFocus
        Exit Function
    End If

    If cboUnit.Text = "" Then
        MsgBox "Fill-up Item Unit.", vbExclamation, "Item Unit Required"
        cboUnit.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function DesExists() As Boolean
    Dim objTempRst  As New ADODB.Recordset
    Dim strsql      As String

    strsql = "select count(*) as the_count from Materials where ItemName = '" & txtDes.Text & "'"
    objTempRst.Open strsql, mmsADOConn, adOpenForwardOnly, , adCmdText
    
    If objTempRst("the_count") > 0 Then
        DesExists = True
    Else
        DesExists = False
    End If

End Function
Private Function GetNextItemID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(ItemID) AS MaxID FROM Materials"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextItemID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextItemID = 1
       Else
           GetNextItemID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function



