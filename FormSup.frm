VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormSuppliers 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUPPLIERS LIBRARY"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15525
   Icon            =   "FormSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10618.23
   ScaleMode       =   0  'User
   ScaleWidth      =   22959.19
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSupPerson 
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
      Left            =   10920
      MaxLength       =   13
      TabIndex        =   4
      Top             =   8640
      Width           =   4440
   End
   Begin VB.TextBox txtSupNum 
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
      Left            =   10920
      MaxLength       =   10
      TabIndex        =   3
      Top             =   8280
      Width           =   4440
   End
   Begin VB.TextBox txtSupName 
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   8280
      Width           =   8085
   End
   Begin VB.TextBox txtSupAd 
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   8640
      Width           =   8085
   End
   Begin VB.TextBox TxtSCode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10800
      MaxLength       =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtSupSearch 
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
      Left            =   4320
      MaxLength       =   50
      TabIndex        =   6
      Top             =   9000
      Width           =   11055
   End
   Begin VB.ComboBox cboSupSearch 
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
      ItemData        =   "FormSup.frx":08CA
      Left            =   1200
      List            =   "FormSup.frx":08D1
      TabIndex        =   5
      Top             =   9000
      Width           =   3045
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1197
      Left            =   0
      TabIndex        =   7
      Top             =   9573
      Width           =   16917
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   150
         Width           =   2300
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
         Left            =   2500
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   150
         Width           =   2500
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
         TabIndex        =   8
         Top             =   150
         Width           =   2300
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
         Left            =   7100
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   2300
      End
   End
   Begin MSComctlLib.ListView lvwSuppliers 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   14208
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "SEARCH"
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
      Left            =   120
      TabIndex        =   19
      Top             =   9120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "CODE"
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
      Index           =   0
      Left            =   11040
      TabIndex        =   14
      Top             =   10080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
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
      Index           =   4
      Left            =   9840
      TabIndex        =   18
      Top             =   8280
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "PERSON"
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
      Left            =   9840
      TabIndex        =   17
      Top             =   8640
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADDRESS"
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
      Left            =   120
      TabIndex        =   16
      Top             =   8640
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "NAME"
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
      Left            =   120
      TabIndex        =   15
      Top             =   8280
      Width           =   555
   End
End
Attribute VB_Name = "FormSuppliers"
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


'Private Const S_Code                     As Long = 1
Private Const S_Address                  As Long = 1
Private Const S_Number                   As Long = 2
Private Const S_Person                   As Long = 3
Private Const S_SupID                    As Long = 4
'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
    Load Me
   ConnectToDB
   SetlvwSuppliers
   LoadSuppliers
   cboSupSearch.Text = "NAME"
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub

'---------------------------------------------------------------------------------
'                                   C O N T R O L S   E V E N T S
'---------------------------------------------------------------------------------
Private Sub txtSupName_GotFocus()
   txtSupName.SelStart = 0
   txtSupName.SelLength = Len(txtSupName.Text)
   cboSupSearch.Text = ""
End Sub
Private Sub TxtSupName_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
    txtSupAd.SetFocus
  End If
End Sub
Private Sub txtSupAd_GotFocus()
   txtSupAd.SelStart = 0
   txtSupAd.SelLength = Len(txtSupAd.Text)
   cboSupSearch.Text = ""
End Sub
Private Sub TxtSupAd_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
    txtSupNum.SetFocus
  End If
End Sub
Private Sub txtSupNum_GotFocus()
   txtSupNum.SelStart = 0
   txtSupNum.SelLength = Len(txtSupNum.Text)
   cboSupSearch.Text = ""
End Sub
Private Sub TxtSupNum_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
    txtSupPerson.SetFocus
  End If
End Sub
Private Sub txtSupPerson_GotFocus()
   txtSupPerson.SelStart = 0
   txtSupPerson.SelLength = Len(txtSupPerson.Text)
   cboSupSearch.Text = ""
End Sub
Private Sub txtSupPerson_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
    cmdSave.SetFocus
  End If
End Sub
'------------------------------------------------------------------------------------
'                  S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboSupSearch_GotFocus()
   'cboSupSearch.Text = "NAME"
   cboSupSearch.SelStart = 0
   cboSupSearch.SelLength = Len(cboSupSearch.Text)
End Sub
Private Sub cboSupSearch_Click()
  'ButtonState False
  'cmdSave.Enabled = False
  lvwSuppliers.Enabled = True
End Sub
Private Sub cboSupSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii < 255 Then
      SendKeys_
   End If
   If KeyAscii = 13 Then
      txtSupSearch.SetFocus
   End If
End Sub
Private Sub txtSupSearch_Change()
Dim strsql       As String
Dim SuppliersLI  As ListItem

On Error GoTo LocalError
    If cboSupSearch.Text = "NAME" Then
       strsql = "Select * from Suppliers where SupName like '" & txtSupSearch.Text & "%'" _
             & "Order by SupName"
    ElseIf cboSupSearch.Text = "ADDRESS" Then
       strsql = "Select * from Suppliers where SupAddress like '" & txtSupSearch.Text & "%'" _
             & "Order by SupAddress"
    Else
       strsql = "Select * from Suppliers where SupName like '" & txtSupSearch.Text & "%'" _
             & "Order by SupName"
    End If
      
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute

    lvwSuppliers.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set SuppliersLI = lvwSuppliers.ListItems.Add(, , !SupName & "")
            SuppliersLI.SubItems(S_Address) = !SupAddress & ""
            SuppliersLI.SubItems(S_Number) = !SupNumber & ""
            SuppliersLI.SubItems(S_Person) = !SupPerson & ""
            SuppliersLI.SubItems(S_SupID) = CStr(!SupId)
            .MoveNext
        Loop
     End With
     
    
    Set SuppliersLI = Nothing
    Set mmsADORst = Nothing

LocalError:
    Exit Sub
End Sub
Private Sub txtSupSearch_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
    lvwSuppliers.SetFocus
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
    txtSupName.SetFocus
    cboSupSearch.Text = ""
End Sub
Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
Private Sub CmdUpdate_Click()


End Sub
Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
Private Sub cmdDelete_Click()
    If lvwSuppliers.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        Exit Sub
    End If

    If MsgBox("Are you sure that you want to delete the item on the list " _
              , vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    ConnectToDB
    mmsAdoCmd.CommandText = "DELETE FROM Suppliers WHERE SupName =  '" & txtSupName.Text & "'"
    mmsAdoCmd.Execute
    ClearBox
    LoadSuppliers
End Sub
Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
   If Not DataValidation Then
      Exit Sub
   End If
      If EncodeMode = "A" Then

         lngIDField = GetNextSupID

         strsql = "INSERT INTO Suppliers (  SupID"
         strsql = strsql & "            , SupName"
         strsql = strsql & "            , SupAddress"
         strsql = strsql & "            , SupNumber"
         strsql = strsql & "            , SupPerson"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtSupName.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtSupAd.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtSupNum.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtSupPerson.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwSuppliers.ListItems.Add(, , txtSupName.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(S_SupID) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwSuppliers.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwSuppliers.SelectedItem.SubItems(S_SupID))

         strsql = "UPDATE Suppliers SET "
         strsql = strsql & "  SupName              = '" & Replace$(txtSupName.Text, "'", "''") & "'"
         strsql = strsql & ", SupAddress           = '" & Replace$(txtSupAd.Text, "'", "''") & "'"
         strsql = strsql & ", SupNumber            = '" & Replace$(txtSupNum.Text, "'", "''") & "'"
         strsql = strsql & ", SupPerson            = '" & Replace$(txtSupPerson.Text, "'", "''") & "'"
         strsql = strsql & " WHERE SupID = " & lngIDField
          
         lvwSuppliers.SelectedItem.Text = txtSupName.Text
         PopulateItem lvwSuppliers.SelectedItem
    End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwSuppliers.Enabled = True
    lvwSuppliers.SetFocus
    ButtonState True
    
    cboSupSearch.Text = ""
    txtSupSearch.Text = ""
    
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwSuppliers.Enabled = True
   cboSupSearch.Text = ""
   txtSupSearch.Text = ""
    
    If lvwSuppliers.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
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
Private Sub cmdExit_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
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
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub

'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub lvwSuppliers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  'LoadMaterials
    With lvwSuppliers
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
        
    If Not lvwSuppliers.SelectedItem Is Nothing Then
        lvwSuppliers.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwSuppliers_ItemClick(ByVal Item As MSComctlLib.ListItem)
      With Item
        txtSupName.Text = .Text
        txtSupAd.Text = .SubItems(S_Address)
        txtSupNum.Text = .SubItems(S_Number)
        txtSupPerson.Text = .SubItems(S_Person)
     End With
End Sub
Private Sub lvwSuppliers_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub
Private Sub lvwSuppliers_DblClick()
    ButtonState False
    BoxState True
    txtSupName.SetFocus
    If EncodeMode = "S" Then
           With lvwSuppliers
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwSuppliers_ItemClick .SelectedItem
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
    
        mstrSQL = "select * from Suppliers"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open mstrSQL, mmsADOConn, adOpenDynamic, adLockOptimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub SetlvwSuppliers()
    With lvwSuppliers
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Name", .Width * 0.32
        .ColumnHeaders.Add , , "Address", .Width * 0.32
        .ColumnHeaders.Add , , "Number", .Width * 0.17
        .ColumnHeaders.Add , , "Reference", .Width * 0.17
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(S_Address) = txtSupAd.Text
        .SubItems(S_Number) = txtSupNum.Text
        .SubItems(S_Person) = txtSupPerson.Text
    End With
End Sub
Private Sub LoadSuppliers()
    Dim strsql        As String
    Dim SuppliersLI  As ListItem
                                 
    strsql = "SELECT SupName" _
           & "     , SupAddress" _
           & "     , SupNumber" _
           & "     , SupPerson" _
           & "     , SupID" _
           & "  FROM Suppliers " _
           & " ORDER BY SupName"

    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwSuppliers.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set SuppliersLI = lvwSuppliers.ListItems.Add(, , !SupName & "")
            SuppliersLI.SubItems(S_Address) = !SupAddress & ""
            SuppliersLI.SubItems(S_Number) = !SupNumber & ""
            SuppliersLI.SubItems(S_Person) = !SupPerson & ""
            SuppliersLI.SubItems(S_SupID) = CStr(!SupId)
            .MoveNext
        Loop
     End With
    
              With lvwSuppliers
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwSuppliers_ItemClick .SelectedItem
                  End If
              End With
     
    Set SuppliersLI = Nothing
    Set mmsADORst = Nothing

End Sub
Private Sub ClearBox()
    txtSupName.Text = ""
    txtSupAd.Text = ""
    txtSupNum.Text = ""
    txtSupPerson.Text = ""
    cboSupSearch.Text = ""
    txtSupSearch.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtSupName.Enabled = boxEnabled
    txtSupAd.Enabled = boxEnabled
    txtSupNum.Enabled = boxEnabled
    txtSupPerson.Enabled = boxEnabled
    cboSupSearch.Enabled = Not boxEnabled
    txtSupSearch.Enabled = Not boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwSuppliers.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdDelete.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub
Private Sub SendKeys_()
    'SendKeys "{left}"
    'SendKeys "{del}"
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtSupName.Text = "" Then
         MsgBox "Fill-up Supplier's Name.", vbExclamation, "Supplier's Name Required"
         txtSupName.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
        If NameExists Then
            MsgBox "Supplier's Name Already Exist", vbExclamation, "Duplicate Name"
            txtSupName.SetFocus
            Exit Function
        End If
     End If
    If txtSupAd.Text = "" Then
        MsgBox "Fill-up Supplier's Address.", vbExclamation, "Address Required"
        txtSupAd.SetFocus
        Exit Function
    End If
    If txtSupNum.Text = "" Then
        MsgBox "Fill-up Supplier's Contact Number.", vbExclamation, "Number Required"
        txtSupNum.SetFocus
        Exit Function
    End If
    If txtSupPerson.Text = "" Then
        MsgBox "Fill-up Supplier's Contact Person.", vbExclamation, "Contact Person Required"
        txtSupPerson.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function NameExists() As Boolean
    Dim objTempRst  As New ADODB.Recordset
    Dim strsql      As String

    strsql = "select count(*) as the_count from Suppliers where SupName = '" & txtSupName.Text & "'"
    objTempRst.Open strsql, mmsADOConn, adOpenForwardOnly, , adCmdText
    
    If objTempRst("the_count") > 0 Then
        NameExists = True
    Else
        NameExists = False
    End If

End Function
Private Function GetNextSupID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(SupID) AS MaxID FROM Suppliers"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextSupID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextSupID = 1
       Else
           GetNextSupID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Public Function ButtonShortcuts()
On Error GoTo LocalError
   
   If ButtonPress = "A" Or ButtonPress = "a" Then
      CmdAdd_Click
    ElseIf ButtonPress = "U" Or ButtonPress = "u" Then
      CmdUpdate_Click
    ElseIf ButtonPress = "D" Or ButtonPress = "d" Then
      cmdDelete_Click
    ElseIf ButtonPress = "X" Or ButtonPress = "x" Then
      CmdExit_Click
    ElseIf ButtonPress = "S" Or ButtonPress = "s" Then
      cboSupSearch.SetFocus
      cboSupSearch.Text = "NAME"
    Exit Function
    End If

LocalError:
    Exit Function
End Function



