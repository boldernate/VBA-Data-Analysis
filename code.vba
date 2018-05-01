Sub Combine_Workbooks()
'*************************************************************************
' Combines the active worksheets of all open workbooks into one new sheet
'*************************************************************************
Set tw = ThisWorkbook.Sheets.Add
tw.Name = "Combined Data" 'Change the name accordingly

For Each w In Workbooks
    If Not w.Name = ThisWorkbook.Name Then
        Set a = w.ActiveSheet
        a.UsedRange.Copy
        tw.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
    End If
Next w

tw.Rows(1).Delete ' Delete Empty row on top

End Sub

Sub Combine_Worksheets()
'*************************************************************************
' Combines all worksheets in the active workbook
'*************************************************************************
Set t = ActiveWorkbook

For Each w In t.Worksheets
  If w.Name = "Combined Sheet" Then
      MsgBox ("Data has already been combined!")
      Exit Sub
  End If
Next w

Set tw = t.Sheets.Add
tw.Name = "Combined Sheet"

Sheet1 = True

For Each w In t.Worksheets
  If Not w.Name = tw.Name Then
    If Sheet1 = True Then
      w.UsedRange.Copy
      tw.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    Else
      With w.UsedRange
        .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0).Copy
      End With
      tw.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
    End If
  End If
End Sub
Sub Vlookup_Columns()
'******************************************************************************
' imports columns using vlookup if the values in column A are an exact match
'******************************************************************************
Set tw = ThisWorkbook.Sheets(1)
Set aw = ActiveSheet

' Adds column headings to Nasdaq Data
tw.Range("G1") = "Header B"
tw.Range("H1") = "Header C"
tw.Range("I1") = "Header D"
tw.Range("J1") = "Header E"
tw.Range("K1") = "Header F"

twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    lv = tw.Cells(I, "A").Value
    Set rv = aw.Range("A2:F" & awlr)
    
    bresult = Application.VLookup(lv, rv, 2, False)
    cresult = Application.VLookup(lv, rv, 3, False)
    dresult = Application.VLookup(lv, rv, 4, False)
    eresult = Application.VLookup(lv, rv, 5, False)
    fresult = Application.VLookup(lv, rv, 6, False)

    tw.Cells(I, "B") = bresult
    tw.Cells(I, "C") = cresult
    tw.Cells(I, "D") = dresult
    tw.Cells(I, "E") = eresult
    tw.Cells(I, "F") = fresult
Next I

End Sub

Sub Condense_Data()
'**********************************************************************
' Uses RemoveDuplicates and SumIfs to condense the data to unique rows
'**********************************************************************
Set a = ActiveWorkbook
Set aw = a.ActiveSheet
Set tw = a.Sheets.Add
tw.Name = "Data Consolidated"

aw.UsedRange.Copy
tw.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

tw.Range("A:B").RemoveDuplicates Columns:=Array(1), Header:=xlYes

twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    Set sumrng = aw.Range("B2:B" & awlr) 'Assumes sum range is in column B
    Set critrng1 = aw.Range("A2:A" & awlr) 'Assumes primary key in column A
    crit1 = tw.Cells(I, "A").Value

    bresult = Application.SumIfs(sumrng, critrng1, crit1)

    tw.Cells(I, "B") = bresult
Next I
tw.Cells(1, 1).Select
End Sub

Sub Copy_Specific_Columns()
'**********************************************************************
' Copies specified columns from a larger worksheet
'**********************************************************************
Set aw = ActiveSheet
Set tw = ThisWorkbook.Sheets.Add
tw.Name = "Specific Columns"

headers = Array("Header1", "Header2", "Header3")

For I = 1 To LBound(headers)
    Set r = Rows(1).Find(headers(I))
    aw.Columns(r.Column).Copy
    tw.Cells(1, Columns.Count).End(xlToLeft).Column.Offset(, 1).PasteSpecial xlPasteValuesAndNumberFormats
Next I
End Sub
Sub Format_Date()
'***********************************************************************
' Formats cells in selection as date
'***********************************************************************

For Each Cell In Selection
    Cell.NumberFormat = "mmm yyyy"
    
Next

End Sub

Sub Remove_Rows_Containing()

'**************************************************************************
' Deletes rows containing the string entered in the input box that pops up
'**************************************************************************

Set tw = ActiveSheet

q = Trim(InputBox("Type the string you want to identify and delete the corresponding rows"))

Set r = tw.Columns(1).Find(q)

Do While Not r Is Nothing
  Set r = tw.Columns(1).Find(q)
  If Not r Is Nothing Then
    tw.Rows(r.Row).Delete
  Else
    Exit Do

  End If
Loop
End Sub
Sub Create_Pivot()
'***********************************************************************
' Creates a basic pivot table that can be easily scaled up or down
'***********************************************************************
Dim PivCache As PivotCache
Dim PivTable As PivotTable


Set dw = ActiveWorkbook.ActiveSheet 'Worksheet containing the data
Set pw = ActiveWorkbook.Sheets.Add
pw.Name = "Pivot Summary"

dwlr = dw.Cells(Rows.Count, "A").End(xlUp).Row
dwlc = dw.Cells(1, Columns.Count).End(xlToLeft).Column

Set prng = dw.Cells(1, 1).Resize(dwlr, dwlc)

' Inserts Pivot Cache needed before making the pivot table
Set PivCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=prng). _
CreatePivotTable(TableDestination:=pw.Cells(2, 2), _
TableName:="BasicPivotTable")

' Inserts Pivot Table
Set PivTable = PivCache.CreatePivotTable _
(TableDestination:=pw.Cells(2, 2), TableName:="BasicPivotTable")

' Inserts Row Fields Below
With pw.PivotTables("BasicPivotTable").PivotFields("Year")
  .Orientation = xlRowField
  .Position = 1
End With

With pw.PivotTables("BasicPivotTable").PivotFields("Month")
  .Orientation = xlRowField
  .Position = 2
End With

' Inserts Column Fields
With pw.PivotTables("BasicPivotTable").PivotFields("Product")
  .Orientation = xlColumnField
  .Position = 1
End With

' Inserts Data Fields
With pw.PivotTables("BasicPivotTable").PivotFields("Amount")
  Orientation = xlDataField
  .Position = 1
  .Function = xlSum
  .NumberFormat = "#,##0"
  .Name = "Revenue "
End With
' Formats the table
pw.PivotTables("BasicPivotTable").ShowTableStyleRowStripes = TrueActiveSheet.PivotTables("BasicPivotTable").TableStyle2 = "PivotStyleMedium9"




End Sub
Sub render_dates_24_months()
' renders eomonth formula to find the last day of each of the last 24 months
Set tw = ActiveSheet

j = 25

For I = 2 To 26

    tw.Cells(I, "A") = Application.EoMonth(Now, "-" & j)
    j = j - 1
Next I

End Sub

Sub Add_Data_Column_Monthly_Data()
Set aw = ActiveSheet
Set tw = ThisWorkbook.Sheets("Monthly Data")

twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column
twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row

awlc = aw.Cells(1, Columns.Count).End(xlToLeft).Column
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    Set r = aw.Columns(1).Find(tw.Cells(I, "B").Value)
    If Not r Is Nothing Then
      tw.Cells(I, twlc + 1) = aw.Cells(rn, awlc).Value
    Else
      tw.Cells(I, twlc + 1) = "Not Found"
    End If
Next I




End Sub

