Attribute VB_Name = "Printing"
'-------------------------------------------------------------------------------
'                  E X C E L     P R I N T
'------------------------------------------------------------------------------
Public Function ListViewPrint()
    Dim ExcelObj   As Object
    Dim ExcelBook  As Object
    Dim ExcelSheet As Object
    Dim lst As ListItem, lst1 As ListSubItem, row As Integer, col As Integer, i As Integer
    Dim AppExcel   As Variant

    Set AppExcel = CreateObject("Excel.application")
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelBook = ExcelObj.WorkBooks.Add
    Set ExcelSheet = ExcelBook.Worksheets(1)
    
    'MsgBox "enable disable scripts"
       
    With ExcelObj.activesheet
          .Pagesetup.Orientation = 2 '--LANDSCAPE (2) xlPortrait
          '.Pagesetup.RightHeader = "Date: " & "&D" & " " & Format(Time, "hh:mm") & " Page: " & "&P of &N"
          .Pagesetup.LeftMargin = 25
          .Pagesetup.RightMargin = 15
          .Pagesetup.TopMargin = 30
          .Pagesetup.BottomMargin = 30
    End With
    With ExcelSheet
          '.Pagesetup.RightFooter = " Page: " & "&P"
          .range("A:L").Font.Size = 9
          .range("A:L").RowHeight = 12
          .range("A:A").Font.Bold = True
          .range("A1").Value = FormReport.lblComp.Caption
          .range("A2").Value = FormMainMenu.lblHeader
          .range("A5").Value = ReportTittle
          .range("A4").Value = ReportType & " REPORT"
          .range("A6").Value = "REPORT PERIOD  :  " & StartMonth & "-" & EndMonth
          
           '-----------------------------------------
           '          M R R   S E T T I N G S
           '-----------------------------------------
            If ReportTransact = "MRR" Then
                 If ReportType = "NUMBER" Then
                    .range("H:H").NumberFormat = "#,##0.00"
                    .Columns.columnwidth = 12: .range("D:G").columnwidth = 25: .Columns(2).columnwidth = 0
                    .range("H:H").Font.Bold = True
                       With ExcelObj.activesheet
                         .Pagesetup.Orientation = 2
                       End With
                 ElseIf ReportType = "SUPPLIER" Then
                   .range("C:C").NumberFormat = "#,##0.00": .range("F:G").NumberFormat = "#,##0.00"
                   .Columns.columnwidth = 2: .Columns(5).columnwidth = 30: .range("C:C").columnwidth = 8:
                   .range("D:D").columnwidth = 5: .range("F:H").columnwidth = 12:
                   .range("H:H").Font.Bold = True
                 End If
            End If
            
           '-----------------------------------------
           '          M I S   S E T T I N G S
           '-----------------------------------------
           If ReportTransact = "MIS" Then
               If ReportType = "NUMBER" Then
                .range("C:C").NumberFormat = "#,##0.00": .range("F:H").NumberFormat = "#,##0.00"
                .Columns.columnwidth = 2: .range("C:C").columnwidth = 8: .range("D:D").columnwidth = 5:
                .range("F:H").columnwidth = 12: .Columns(5).columnwidth = 30
                .range("H:H").Font.Bold = True
               ElseIf ReportType = "CHARGED" Then
                .range("E:E").NumberFormat = "#,##0.00":
                .Columns.columnwidth = 15
                .range("E:E").Font.Bold = True
               ElseIf ReportType = "DEPARTMENT" Or ReportType = "BLOCK" Or ReportType = "RECEIVED" Then
                .range("G:J").NumberFormat = "#,##0.00"
                .Columns.columnwidth = 11: .range("A:B").columnwidth = 2: .range("D:D").columnwidth = 9
                .range("E:E").columnwidth = 15: .range("F:G").columnwidth = 5
                .range("J:J").Font.Bold = True
               ElseIf ReportType = "MATGROUP" Or ReportType = "BLOCK_SUM" Then
                .range("A:L").Font.Size = 11
                .range("C:D").NumberFormat = "#,##0.00"
                .Columns.columnwidth = 5: .range("B:B").columnwidth = 20: .range("C:D").columnwidth = 15
                .range("D:D").Font.Bold = True
               End If
         End If
         
           '---------------------------------------------
           '          I N V E N T O R Y   S E T T I N G S
           '--------------------------------------------
           If ReportTransact = "INVENTORY" Then
                .range("C:F").NumberFormat = "#,##0.00"
                .Columns.columnwidth = 7
                .Columns(2).columnwidth = 35
                .Columns(7).columnwidth = 25
                .Columns(8).columnwidth = 0
                .range("E:F").columnwidth = 13
          End If
                     '---------------------------------------------
           '          I N V E N T O R Y   S E T T I N G S
           '--------------------------------------------
           If ReportTransact = "FUEL" Then
                .range("C:E").NumberFormat = "#,##0.00"
                .range("D:D").Font.Bold = True
                .Columns.columnwidth = 15
          End If
         
    End With
    
    row = 8
    col = 1
    
    If ReportType = "HISTORY" Then     ' HISTORY LISTVIEW TO EXCEL
            With ExcelSheet
             .Columns.columnwidth = 9
             .Columns(1).columnwidth = 8
             .Columns(2).columnwidth = 11
             .Columns(3).columnwidth = 15
             .range("A6:C6").mergecells = True
             .range("D6:F6").mergecells = True
             .range("G6:I6").mergecells = True
             .range("J6:L6").mergecells = True
             .range("A6:L6").Font.Bold = True
             .range("A6:L6").HorizontalAlignment = -4108   ' ALIGN CENTER
             .range("A7:L7").HorizontalAlignment = -4108
             .range("A6").Value = "--- ITEM REFERENCE ---"
             .range("D6").Value = "--- RECEIVED ---"
             .range("G6").Value = "--- ISSUANCES ---"
             .range("J6").Value = "--- BALANCE ---"
            End With
        row = 8
        col = 1
       For i = col To lvwTransact.ColumnHeaders.Count
        ExcelSheet.cells(row, col) = lvwTransact.ColumnHeaders(i)
        col = col + 1
       Next
        row = 11
        col = 1
        For Each lst In lvwTransact.ListItems
          col = 1
          ExcelSheet.cells(row, col) = lst.Text
          col = col + 1
          For Each lst1 In lst.ListSubItems
            ExcelSheet.cells(row, col) = lst1.Text
            col = col + 1
          Next
          row = row + 1
        Next
       ExcelSheet.range("A3").Value = "ITEM NAME : " & ItemClick
    
    Else ' SUMMARY LISTVIEW TO EXCEL
        
      With FormReport
      For i = 1 To .lvwSummary.ColumnHeaders.Count
        ExcelSheet.cells(row, col) = .lvwSummary.ColumnHeaders(i)
        col = col + 1
      Next
       row = 9
       col = 1
        For Each lst In .lvwSummary.ListItems
          col = 1
          ExcelSheet.cells(row, col) = lst.Text
          col = col + 1
          For Each lst1 In lst.ListSubItems
            ExcelSheet.cells(row, col) = lst1.Text
            col = col + 1
          Next
          row = row + 1
        Next
        End With
        
    End If
    
    ExcelObj.Visible = True
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
End Function
