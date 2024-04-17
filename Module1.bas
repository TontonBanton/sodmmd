Attribute VB_Name = "SQLs"
    Public Function ConvertUpper(pintKeyValue As Integer) As Integer
'  Common function to force alphabetic keyboard characters to uppercase
'  when called from the KeyPress event.
'  Typical call:
'      KeyAscii = ConvertUpper(KeyAscii)
    If Chr$(pintKeyValue) >= "a" And Chr$(pintKeyValue) <= "z" Then
        pintKeyValue = pintKeyValue - 32
    End If
    ConvertUpper = pintKeyValue
End Function
'-----------------------------------------------------------------------------Private Sub ConnectToDB()
Public Function ConnectToDB()
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

End Function
Public Function POSaveTemp() As String
With FormPO
POSaveTemp = "INSERT INTO PODetailsTemp (POID, POItemNo, PONum, PODate, POPrs, POArea, POSupplier, POTerms, POWork "
    POSaveTemp = POSaveTemp & ", POEquip, PORemark, POGroup, POCode, POItem, POUnit, POQty, POCost, POAmount, POStatus, POTr, POMrr"
    POSaveTemp = POSaveTemp & " ) VALUES ("
    POSaveTemp = POSaveTemp & .POId
    POSaveTemp = POSaveTemp & ", '" & .POItemNo & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPONum.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPODate.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOPrs.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOArea.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOSupplier.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOTerms.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOWork.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOEquip.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPORemark.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & .ItemGroup & "'"
    POSaveTemp = POSaveTemp & ", '" & .ItemCode & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOItem.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOUnit.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOQty.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOCost.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOAmount.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '" & Replace$(.txtPOStatus.Text, "'", "''") & "'"
    POSaveTemp = POSaveTemp & ", '-'"
    POSaveTemp = POSaveTemp & ", '-' )"
    
End With
End Function

'-----------------------------------------------------------
Public Function AssetInsert() As String
      AssetInsert = "INSERT INTO Asset    (  AssetId"
         strsql = strsql & "            , AssetName"
         strsql = strsql & "            , AssetGroup"
         strsql = strsql & "            , Unit"
         strsql = strsql & "            , Cost"
         strsql = strsql & "            , Avail"
         strsql = strsql & "            , AssetAmount"
         strsql = strsql & "            , ATFNum"
         strsql = strsql & "            , ATFDate"
         strsql = strsql & "            , AssetTag"
         strsql = strsql & "            , SerialNum"
         strsql = strsql & "            , PONum"
         strsql = strsql & "            , Status"
         strsql = strsql & "            , InPosition"
         strsql = strsql & "            , BorrowDate"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & ItemInvID
         strsql = strsql & ", '" & Replace$(txtAssetName.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetGroup.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetCost.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetAvail.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetAmount.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetATF.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtATFDate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetTag.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetSerial.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtAssetPO.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$("IN", "'", "''") & "'"
         strsql = strsql & ", '" & Replace$("WAREHOUSE", "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtATFDate.Text, "'", "''") & "'"
         strsql = strsql & ")"
End Function
