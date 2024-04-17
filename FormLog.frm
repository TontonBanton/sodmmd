VERSION 5.00
Begin VB.Form FormLog 
   Caption         =   "SOUTH DAVAO DEVELOPMENT CO. INC."
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   Icon            =   "FormLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleMode       =   0  'User
   ScaleWidth      =   25000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEncodeMode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4665
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   2200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2350
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   2200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtPWord 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1110
      End
   End
End
Attribute VB_Name = "FormLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn              As ADODB.Connection
Private mmsAdoCmd               As ADODB.Command
Private mmsADORst               As ADODB.Recordset
Private strsql                  As String
Dim User, PWord                 As String
Private Sub Form_Load()
   ConnectToDB
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtPWord.SetFocus
   End If
End Sub
Private Sub txtUser_GotFocus()
  txtUser.Text = ""
End Sub
Private Sub txtPWord_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
     cmdOk.SetFocus
   End If
End Sub
Private Sub txtPWord_GotFocus()
  txtPWord.Text = ""
End Sub
Private Sub cmdOk_Click()
    User = txtUser.Text
    PWord = txtPWord.Text
    strsql = "Select * from PWord Where UName like '" & User & "' and PWord like '" & PWord & "'"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    If mmsADORst.EOF = True Then
       MsgBox "INVALID ENTRY", vbCritical, "Security"
       txtUser.Text = ""
       txtPWord.Text = ""
       txtUser.SetFocus
    Else
       txtUser.Text = ""
       txtPWord.Text = ""
       If txtEncodeMode.Text = "SETTINGS" Then
            FormLog.Hide
            FormSettings.Show
       End If
       Unload Me
       FormMainMenu.Show
    End If
    
   Set mmsADORst = Nothing
End Sub
Private Sub cmdCancel_Click()
If txtEncodeMode.Text = "ADJUST" Or txtEncodeMode.Text = "SETTINGS" Then
    Unload Me
Else
    txtUser.Text = ""
    txtPWord.Text = ""
    txtUser.SetFocus
End If
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

Private Sub cmdOk_GotFocus()
   cmdOk.BackColor = &HC0FFC0
End Sub
Private Sub cmdOk_LostFocus()
   cmdOk.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
