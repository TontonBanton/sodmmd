Attribute VB_Name = "ElementState"
Public Sub GetCurrentDate()
    If (Month(Now)) = 1 Then
         FormReport.cboMonth.Text = "DECEMBER"
    Else
         FormReport.cboMonth.Text = UCase(MonthName(Month(Now) - 1))
    End If
       
    If (Month(Now)) = 1 Then
        FormReport.txtYear.Text = ((Year(Now) - 1))
    Else
        FormReport.txtYear.Text = Year(Now)
    End If
    
End Sub
Public Sub ReportRange()
On Error GoTo LocalError
Dim LastDay As Long
With FormReport
   .MonthSummary = Month(CDate("1 " & .cboMonth.Text))
   .StartMonth = .MonthSummary & "/" & "1" & "/" & .txtYear.Text
   LastDay = Day(DateSerial(Year(.txtYear.Text), Month(.StartMonth) + 1, 0))
   .EndMonth = .MonthSummary & "/" & LastDay & "/" & .txtYear.Text
   .txtDRStart.Text = Format$(.StartMonth, "mm/dd/yyyy")
   .txtDREnding.Text = Format$(.EndMonth, "mm/dd/yyyy")
End With
LocalError: Exit Sub
End Sub
Public Sub ClearOptions()
With FormReport
   .OptMRRNumber.Value = False
   .OptMRRSupplier.Value = False
   .OptMISNumber.Value = False
   .OptMISCCC.Value = False
   .OptMISDept.Value = False
   .OptMISCharge.Value = False
   .OptInventory.Value = False
   .OptFuel.Value = False
   .OptAssets.Value = False
End With
End Sub
Public Sub BoxState(boxEnabled As Boolean)
With FormPO
    .txtPODate.Enabled = boxEnabled
    .txtPOArea.Enabled = boxEnabled
    .txtPOPrs.Enabled = boxEnabled
    .txtPOSupplier.Enabled = boxEnabled
    .txtPOTerms.Enabled = boxEnabled
    .txtPOWork.Enabled = boxEnabled
    .txtPOEquip.Enabled = boxEnabled
    .txtPORemark.Enabled = boxEnabled
End With
End Sub
Public Sub ItemBoxState(boxEnabled As Boolean)
With FormPO
    .txtPOItem.Enabled = boxEnabled
    .txtPOQty.Enabled = boxEnabled
    .txtPOUnit.Enabled = boxEnabled
    .txtPOCost.Enabled = boxEnabled
End With
End Sub
Public Sub ButtonState(buttonEnabled As Boolean)
With FormPO
    .lvwPO.Enabled = buttonEnabled
    .cmdNew.Enabled = buttonEnabled
    .cmdAdd.Enabled = buttonEnabled
    .cmdSearch.Enabled = buttonEnabled
    .cmdPrint.Enabled = buttonEnabled
    .cmdCancel.Enabled = buttonEnabled
    '.cmdEdit.Enabled = buttonEnabled
    '.cmdDelete.Enabled = buttonEnabled
End With
End Sub
Public Sub ClearBox()
With FormPO
    .txtPOWork2.Visible = False
    .lvwPO.ListItems.Clear: .txtPOTotal.Text = "": '.txtPOSupAd.Text = ""
    .txtPONum.Text = "": .txtPODate.Text = "__/__/____": .txtPOPrs.Text = ""
    .txtPOSupplier.Text = "": .txtPOTerms.Text = "COD"
    .txtPOWork.Clear: .txtPOEquip.Clear: .txtPORemark.Text = "": .txtPOStatus.Text = "-"
End With
End Sub
Public Sub ClearItemBox()
With FormPO
    .ItemCode = "": .ItemGroup = "": .txtPOTotal.Text = ""
    .txtPOItem.Text = "": .txtPOGroup.Text = "": .txtPOUnit.Text = ""
    .txtPOQty.Text = "": .txtPOCost.Text = "": .txtPOAmount.Text = ""
End With
End Sub
Public Sub ClearFrame()
With FormPO
    .frameSuppliers.Visible = False: .frameItemAdd.Visible = False
    .frameItems.Visible = False: .frameSearch.Visible = False
End With
End Sub
Public Function NumWords()
Dim Total, Cents, Amount As String
Dim NumStr As Integer
    totals = CStr(FormPO.txtPOTotal.Text)
    Cents = Right(totals, 2)
    NumStr = Len(totals) - 3
    Amount = Left(totals, NumStr): FormPO.txtNumWords.Text = Amount & " " & Cents & "/100"
    
End Function
