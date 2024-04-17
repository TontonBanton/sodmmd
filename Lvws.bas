Attribute VB_Name = "Lvws"
Public Function SetlvwPOMain()
With FormPO.lvwPO
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , " ", .Width * 0.05
    .ColumnHeaders.Add , , "CODE", .Width * 0#
    .ColumnHeaders.Add , , "ITEM", .Width * 0.35
    .ColumnHeaders.Add , , "QTY", .Width * 0.13
    .ColumnHeaders.Add , , "UNIT", .Width * 0.07
    .ColumnHeaders.Add , , "U/P", .Width * 0.15
    .ColumnHeaders.Add , , "AMOUNT", .Width * 0.2
    .ColumnHeaders.Item(4).Alignment = lvwColumnRight
    .ColumnHeaders.Item(6).Alignment = lvwColumnRight
    .ColumnHeaders.Item(7).Alignment = lvwColumnRight
End With
End Function
Public Function SetLvwPoSuppliers()
With FormPO.frameSuppliers
    .Top = 850: .Left = 10300: .Width = 14850: .Height = 8630: .Visible = True
End With
With FormPO.lvwPOSuppliers
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Name", .Width * 0.95
    .ColumnHeaders.Add , , "Id", .Width * 0#
End With
FormPO.txtSupSearch.SetFocus
End Function
Public Function SetLvwPoItems()
With FormPO.frameItems
    .Top = FormPO.Frame1.Top: .Left = FormPO.lvwPO.Left: .Width = FormPO.lvwPO.Width: .Height = 8900: .Visible = True:
End With
With FormPO.lvwPOItems
    .Height = FormPO.frameItems.Height - 800: .Top = 950
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "GROUP", .Width * 0#
    .ColumnHeaders.Add , , "CODE", .Width * 0.1
    .ColumnHeaders.Add , , "ITEM NAME", .Width * 0.6
    .ColumnHeaders.Add , , "STOCK", .Width * 0#
    .ColumnHeaders.Add , , " ", .Width * 0.08
    .ColumnHeaders.Add , , "ID", .Width * 0#
    .ColumnHeaders.Item(4).Alignment = lvwColumnRight
End With
FormPO.txtItemSearch.SetFocus
End Function
Public Function SetLvwPoSearch()
With FormPO.frameSearch
    .Top = FormPO.lvwPO.Top - 100: .Left = FormPO.lvwPO.Left: .Height = FormPO.lvwPO.Height + 100: .Width = FormPO.lvwPO.Width: .Visible = True
End With
With FormPO.lvwPOSearch
    .Height = FormPO.frameSearch.Height - 900
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "NUMBER", .Width * 0.1
    .ColumnHeaders.Add , , "DATE", .Width * 0.1
    .ColumnHeaders.Add , , "PRS", .Width * 0.15
    .ColumnHeaders.Add , , "SUPPLIER", .Width * 0.25:
    .ColumnHeaders.Add , , " ", .Width * 0.25
    .ColumnHeaders.Add , , " ", .Width * 0#
    .ColumnHeaders.Add , , " ", .Width * 0#
    .ColumnHeaders.Add , , " ", .Width * 0.1
    .ColumnHeaders.Item(8).Alignment = lvwColumnRight
End With
End Function
'----------------------- REPORT---------------------------
Public Sub SetlvwPO()
    With FormReport.lvwSummary
        .ColumnHeaders.Clear: .ListItems.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.05
        .ColumnHeaders.Add , , " ", .Width * 0.06
        .ColumnHeaders.Add , , " ", .Width * 0.15
        .ColumnHeaders.Add , , " ", .Width * 0.05
        .ColumnHeaders.Add , , " ", .Width * 0.1
        .ColumnHeaders.Add , , "", .Width * 0.15
        .ColumnHeaders.Add , , "", .Width * 0.1
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
    
    If FormReport.ReportType = "NUMBER" Then
      .ColumnHeaders.Item(8).Alignment = lvwColumnRight
    End If
    If FormReport.ReportType = "AREA" Then
       .ColumnHeaders.Item(9).Alignment = lvwColumnRight
    End If
    End With
End Sub
Public Sub SetlvwPO1()
    With FormReport.lvwSummary
        .ColumnHeaders.Clear: .ListItems.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.08
        .ColumnHeaders.Add , , " ", .Width * 0.08
        .ColumnHeaders.Add , , " ", .Width * 0.17
        .ColumnHeaders.Add , , " ", .Width * 0.05
        .ColumnHeaders.Add , , " ", .Width * 0.05
        .ColumnHeaders.Add , , "", .Width * 0.17
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Item(4).Alignment = lvwColumnRight: .ColumnHeaders.Item(7).Alignment = lvwColumnRight
        .ColumnHeaders.Item(8).Alignment = lvwColumnRight: .ColumnHeaders.Item(9).Alignment = lvwColumnRight
    End With
End Sub
Public Sub SetlvwPO2()
    With FormReport.lvwSummary
        .ColumnHeaders.Clear: .ListItems.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.2
        .ColumnHeaders.Add , , " ", .Width * 0.15
        .ColumnHeaders.Item(2).Alignment = lvwColumnRight
    End With
End Sub
Public Sub SetlvwPODetails()
    With FormReport.lvwSummary
        .ColumnHeaders.Clear: .ListItems.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.15
        .ColumnHeaders.Add , , "B1", .Width * 0.12
        .ColumnHeaders.Add , , "B2", .Width * 0.12
        .ColumnHeaders.Add , , "B3", .Width * 0.12
        .ColumnHeaders.Add , , "B4", .Width * 0.12
        .ColumnHeaders.Add , , "DCO", .Width * 0.12
        .ColumnHeaders.Add , , "PFI", .Width * 0.12
        .ColumnHeaders.Add , , "TOTAL", .Width * 0.12
    End With
End Sub
Public Sub SetlvwTRPO()
    With FormTransfer.framePO
        .Top = 2280: .Left = 2840: .Height = 7300: .Width = 19000: .Visible = True
    End With
    With FormTransfer.lvwPO
        .Height = FormTransfer.framePO.Height - 800
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.1
        .ColumnHeaders.Add , , "", .Width * 0.35
        .ColumnHeaders.Add , , "", .Width * 0.15
        .ColumnHeaders.Add , , "", .Width * 0.08
        .ColumnHeaders.Add , , "", .Width * 0.13
        .ColumnHeaders.Item(6).Alignment = lvwColumnRight
    End With
    FormTransfer.txtPOSearch.SetFocus
End Sub
