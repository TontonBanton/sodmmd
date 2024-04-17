VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormTransact 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSACTION LIBRARY"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10428.1
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTransactCode 
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
      Left            =   360
      MaxLength       =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAddTransact 
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
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   8880
      Width           =   8212
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1097
      Left            =   0
      TabIndex        =   2
      Top             =   9600
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
         TabIndex        =   6
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
   End
   Begin MSComctlLib.ListView lvwTransact 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
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
Attribute VB_Name = "FormTransact"
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
   SetlvwTransact
   LoadTransact
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub
Private Sub txtAddTransact_GotFocus()
   txtAddTransact.SelStart = 0
   txtAddTransact.SelLength = Len(txtAddTransact.Text)
End Sub

Private Sub txtAddTransact_KeyPress(KeyAscii As Integer)
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
    txtAddTransact.SetFocus
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
   If Not DataValidation Then
      Exit Sub
   End If
      If EncodeMode = "A" Then

         lngIDField = GetNextTransactID
         
         strsql = "INSERT INTO Transact (  TransactID"
         strsql = strsql & "            , TransactName"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtAddTransact.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwTransact.ListItems.Add(, , txtAddTransact.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(1) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwTransact.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwTransact.SelectedItem.SubItems(1))

         strsql = "UPDATE Transact SET "
         strsql = strsql & "  TransactName       = '" & Replace$(txtAddTransact.Text, "'", "''") & "'"
         strsql = strsql & " WHERE TransactID = " & lngIDField
 
         lvwTransact.SelectedItem.Text = txtAddTransact.Text
         PopulateItem lvwTransact.SelectedItem
      End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwTransact.Enabled = True
    lvwTransact.SetFocus
    ButtonState True
    
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
 txtAddTransact.Text = ""
 cmdAdd.Enabled = True
 cmdExit.Enabled = True
 lvwTransact.Enabled = True
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
Private Sub SetlvwTransact()
    With lvwTransact
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", .Width * 0.93
        .ColumnHeaders.Add , , "", Width * 0#
    End With
End Sub
Private Sub lvwTransact_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'LoadMaterials
    With lvwTransact
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
        
    If Not lvwTransact.SelectedItem Is Nothing Then
        lvwTransact.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwTransact_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtAddTransact.Text = .Text
        txtTransactCode.Text = .SubItems(1)
     End With
End Sub
Private Sub lvwTransact_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub
Private Sub lvwTransact_DblClick()
    ButtonState False
    BoxState True
    'txtProduct.Enabled = False
    'txtClassification.SetFocus
    If EncodeMode = "S" Then
           With lvwTransact
           If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwTransact_ItemClick .SelectedItem
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
Private Sub LoadTransact()
    Dim strsql        As String
    Dim TransactLI  As ListItem
                                 
    strsql = "SELECT * FROM Transact ORDER BY TransactName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    lvwTransact.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set TransactLI = lvwTransact.ListItems.Add(, , !TransactName & "")
            TransactLI.SubItems(1) = !TransactID & ""
            .MoveNext
        Loop
     End With
    
              With lvwTransact
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     'lvwTransact_ItemClick .SelectedItem
                  End If
              End With
     
    Set TransactLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = txtTransactCode.Text
    End With
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtAddTransact.Text = "" Then
         MsgBox "FILL UP Transact", vbExclamation, "Transact"
         txtAddTransact.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function GetNextTransactID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(TransactID) AS MaxID FROM Transact"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextTransactID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextTransactID = 1
       Else
           GetNextTransactID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub ClearBox()
    txtAddTransact.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtAddTransact.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwTransact.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub



