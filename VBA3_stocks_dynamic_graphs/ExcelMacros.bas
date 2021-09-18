Option Explicit

Sub ExposureMacro()

'Declare variables

Dim wb As Workbook
Dim input_sheet, combined_sheet, graph_sheet As Worksheet

Dim total_long, total_short, total_gross, total_net, _
    gross_tech, gross_consume, gross_comms, gross_other, _
    gross_tech_a, gross_consume_a, gross_comms_a, gross_other_a, _
    region_europe, region_us, region_other, _
    region_europe_a, region_us_a, region_other_a, _
    market_large, market_mid, market_small, _
    market_large_a, market_mid_a, market_small_a As Double
    
Dim input_date As Date

Dim rw As Variant

Dim flag As Boolean

Dim new_row As ListRow

Dim response As Integer


'Setting workbook and worksheet variables

Set wb = ActiveWorkbook
Set input_sheet = wb.Sheets("Input")
Set combined_sheet = wb.Sheets("Combined Tables")
Set graph_sheet = wb.Sheets("Graph Tables")


'Check if input_date is given

If IsDate(input_sheet.Range("E4").Value) And Not IsEmpty(input_sheet.Range("E4").Value) Then
    input_date = DateTime.DateSerial(Year(input_sheet.Range("E4").Value), Month(input_sheet.Range("E4").Value), 1)
Else
    MsgBox "Please check Input Date value", , "ALERT"
    Exit Sub
End If


'Check and gather inputs

'Total Exposure table

If IsNumeric(input_sheet.Range("B4").Value) And Not IsEmpty(input_sheet.Range("B4").Value) Then
    total_long = input_sheet.Range("B4").Value
Else
    MsgBox "Please check cell B4", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B5").Value) And Not IsEmpty(input_sheet.Range("B5").Value) Then
    total_short = input_sheet.Range("B5").Value
Else
    MsgBox "Please check cell B5", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B6").Value) And Not IsEmpty(input_sheet.Range("B6").Value) Then
    total_gross = input_sheet.Range("B6").Value
Else
    MsgBox "Please check cell B6", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B7").Value) And Not IsEmpty(input_sheet.Range("B7").Value) Then
    total_net = input_sheet.Range("B7").Value
Else
    MsgBox "Please check cell B7", , "ALERT"
    Exit Sub
End If

'Gross Exposure by Sector table

If IsNumeric(input_sheet.Range("B10").Value) And Not IsEmpty(input_sheet.Range("B10").Value) Then
    gross_tech = input_sheet.Range("B10").Value
Else
    MsgBox "Please check cell B10", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B11").Value) And Not IsEmpty(input_sheet.Range("B11").Value) Then
    gross_consume = input_sheet.Range("B11").Value
Else
    MsgBox "Please check cell B11", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B12").Value) And Not IsEmpty(input_sheet.Range("B12").Value) Then
    gross_comms = input_sheet.Range("B12").Value
Else
    MsgBox "Please check cell B12", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B13").Value) And Not IsEmpty(input_sheet.Range("B13").Value) Then
    gross_other = input_sheet.Range("B13").Value
Else
    MsgBox "Please check cell B13", , "ALERT"
    Exit Sub
End If

'Gross Exposure by Sector table Attribution

If IsNumeric(input_sheet.Range("C10").Value) And Not IsEmpty(input_sheet.Range("C10").Value) Then
    gross_tech_a = input_sheet.Range("C10").Value
Else
    MsgBox "Please check cell C10", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C11").Value) And Not IsEmpty(input_sheet.Range("C11").Value) Then
    gross_consume_a = input_sheet.Range("C11").Value
Else
    MsgBox "Please check cell C11", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C12").Value) And Not IsEmpty(input_sheet.Range("C12").Value) Then
    gross_comms_a = input_sheet.Range("C12").Value
Else
    MsgBox "Please check cell C12", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C13").Value) And Not IsEmpty(input_sheet.Range("C13").Value) Then
    gross_other_a = input_sheet.Range("C13").Value
Else
    MsgBox "Please check cell C13", , "ALERT"
    Exit Sub
End If

'Exposure by Region table

If IsNumeric(input_sheet.Range("B16").Value) And Not IsEmpty(input_sheet.Range("B16").Value) Then
    region_europe = input_sheet.Range("B16").Value
Else
    MsgBox "Please check cell B16", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B17").Value) And Not IsEmpty(input_sheet.Range("B17").Value) Then
    region_us = input_sheet.Range("B17").Value
Else
    MsgBox "Please check cell B17", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B18").Value) And Not IsEmpty(input_sheet.Range("B18").Value) Then
    region_other = input_sheet.Range("B18").Value
Else
    MsgBox "Please check cell B18", , "ALERT"
    Exit Sub
End If

'Exposure by Region table Attribution

If IsNumeric(input_sheet.Range("C16").Value) And Not IsEmpty(input_sheet.Range("C16").Value) Then
    region_europe_a = input_sheet.Range("C16").Value
Else
    MsgBox "Please check cell C16", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C17").Value) And Not IsEmpty(input_sheet.Range("C17").Value) Then
    region_us_a = input_sheet.Range("C17").Value
Else
    MsgBox "Please check cell C17", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C18").Value) And Not IsEmpty(input_sheet.Range("C18").Value) Then
    region_other_a = input_sheet.Range("C18").Value
Else
    MsgBox "Please check cell C18", , "ALERT"
    Exit Sub
End If

'Market Capital Exposure table

If IsNumeric(input_sheet.Range("B21").Value) And Not IsEmpty(input_sheet.Range("B21").Value) Then
    market_large = input_sheet.Range("B21").Value
Else
    MsgBox "Please check cell B21", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B22").Value) And Not IsEmpty(input_sheet.Range("B22").Value) Then
    market_mid = input_sheet.Range("B22").Value
Else
    MsgBox "Please check cell B22", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("B23").Value) And Not IsEmpty(input_sheet.Range("B23").Value) Then
    market_small = input_sheet.Range("B23").Value
Else
    MsgBox "Please check cell B23", , "ALERT"
    Exit Sub
End If

'Market Capital Exposure table Attribution

If IsNumeric(input_sheet.Range("C21").Value) And Not IsEmpty(input_sheet.Range("C21").Value) Then
    market_large_a = input_sheet.Range("C21").Value
Else
    MsgBox "Please check cell C21", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C22").Value) And Not IsEmpty(input_sheet.Range("C22").Value) Then
    market_mid_a = input_sheet.Range("C22").Value
Else
    MsgBox "Please check cell C22", , "ALERT"
    Exit Sub
End If

If IsNumeric(input_sheet.Range("C23").Value) And Not IsEmpty(input_sheet.Range("C23").Value) Then
    market_small_a = input_sheet.Range("C23").Value
Else
    MsgBox "Please check cell C23", , "ALERT"
    Exit Sub
End If


'Override Warning
For Each rw In combined_sheet.Range("TotalExposure[Date]")
    If rw = input_date Then
        response = MsgBox("Data for " & Format(input_date, "m/yyyy") & " already exists. Are you sure you would like to override the data for " & Format(input_date, "m/yyyy") & "?", vbYesNo, "WARNING!")
        
        If response = vbNo Then
            MsgBox "Macro Canceled", , "ALERT"
            Exit Sub
        End If
    End If
    
    Next rw

rw = CDate(1 / 1 / 1900)

'Populate Combined

flag = False

'Total Exposure
For Each rw In combined_sheet.Range("TotalExposure[Date]")
    If rw = input_date Then
        combined_sheet.Cells(rw.Row, 2) = total_long
        combined_sheet.Cells(rw.Row, 3) = total_short
        combined_sheet.Cells(rw.Row, 4) = total_gross
        combined_sheet.Cells(rw.Row, 5) = total_net
        
        flag = True
        
        Exit For
    End If
    
    Next rw

If flag = False Then
    Set new_row = combined_sheet.ListObjects("TotalExposure").ListRows.Add
    With new_row
        .Range(1) = input_date
        .Range(2) = total_long
        .Range(3) = total_short
        .Range(4) = total_gross
        .Range(5) = total_net
    End With
Else
    flag = False
End If

rw = CDate(1 / 1 / 1900)

'Gross Exposure
For Each rw In combined_sheet.Range("GrossExposure[Date]")
    If rw = input_date Then
        combined_sheet.Cells(rw.Row, 8) = gross_tech
        combined_sheet.Cells(rw.Row, 9) = gross_consume
        combined_sheet.Cells(rw.Row, 10) = gross_comms
        combined_sheet.Cells(rw.Row, 11) = gross_other
        combined_sheet.Cells(rw.Row, 12) = gross_tech_a
        combined_sheet.Cells(rw.Row, 13) = gross_consume_a
        combined_sheet.Cells(rw.Row, 14) = gross_comms_a
        combined_sheet.Cells(rw.Row, 15) = gross_other_a
        
        flag = True
        
        Exit For
    End If
    
    Next rw

If flag = False Then
    Set new_row = combined_sheet.ListObjects("GrossExposure").ListRows.Add
    With new_row
        .Range(1) = input_date
        .Range(2) = gross_tech
        .Range(3) = gross_consume
        .Range(4) = gross_comms
        .Range(5) = gross_other
        .Range(6) = gross_tech_a
        .Range(7) = gross_consume_a
        .Range(8) = gross_comms_a
        .Range(9) = gross_other_a
    End With
Else
    flag = False
End If

rw = CDate(1 / 1 / 1900)

'Region Exposure

For Each rw In combined_sheet.Range("RegionExposure[Date]")
    If rw = input_date Then
        combined_sheet.Cells(rw.Row, 18) = region_europe
        combined_sheet.Cells(rw.Row, 19) = region_us
        combined_sheet.Cells(rw.Row, 20) = region_other
        combined_sheet.Cells(rw.Row, 21) = region_europe_a
        combined_sheet.Cells(rw.Row, 22) = region_us_a
        combined_sheet.Cells(rw.Row, 23) = region_other_a
        
        flag = True
        
        Exit For
    End If
    
    Next rw

If flag = False Then
    Set new_row = combined_sheet.ListObjects("RegionExposure").ListRows.Add
    With new_row
        .Range(1) = input_date
        .Range(2) = region_europe
        .Range(3) = region_us
        .Range(4) = region_other
        .Range(5) = region_europe_a
        .Range(6) = region_us_a
        .Range(7) = region_other_a
    End With
Else
    flag = False
End If

rw = CDate(1 / 1 / 1900)

'Market Exposure

For Each rw In combined_sheet.Range("MarketExposure[Date]")
    If rw = input_date Then
        combined_sheet.Cells(rw.Row, 27) = market_large
        combined_sheet.Cells(rw.Row, 28) = market_mid
        combined_sheet.Cells(rw.Row, 29) = market_small
        combined_sheet.Cells(rw.Row, 30) = market_large_a
        combined_sheet.Cells(rw.Row, 31) = market_mid_a
        combined_sheet.Cells(rw.Row, 32) = market_small_a

        flag = True
        
        Exit For
    End If
    
    Next rw

If flag = False Then
    Set new_row = combined_sheet.ListObjects("MarketExposure").ListRows.Add
    With new_row
        .Range(1) = input_date
        .Range(2) = market_large
        .Range(3) = market_mid
        .Range(4) = market_small
        .Range(5) = market_large_a
        .Range(6) = market_mid_a
        .Range(7) = market_small_a
    End With
Else
    flag = False
End If

rw = CDate(1 / 1 / 1900)


'Sort Tables

With combined_sheet.ListObjects("TotalExposure").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("TotalExposure[Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
    .Header = xlYes
    .Apply
End With

With combined_sheet.ListObjects("GrossExposure").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("GrossExposure[Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
    .Header = xlYes
    .Apply
End With

With combined_sheet.ListObjects("RegionExposure").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("RegionExposure[Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
    .Header = xlYes
    .Apply
End With

With combined_sheet.ListObjects("MarketExposure").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("MarketExposure[Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
    .Header = xlYes
    .Apply
End With


'Format date columns

combined_sheet.Range("TotalExposure[Date]").NumberFormat = "m/yyyy"
combined_sheet.Range("GrossExposure[Date]").NumberFormat = "m/yyyy"
combined_sheet.Range("RegionExposure[Date]").NumberFormat = "m/yyyy"
combined_sheet.Range("MarketExposure[Date]").NumberFormat = "m/yyyy"

'Format percent columns

combined_sheet.Range("TotalExposure[[Long]:[Net]]").NumberFormat = "0.0%"
combined_sheet.Range("GrossExposure[[Technology]:[Other]]").NumberFormat = "0.0%"
combined_sheet.Range("RegionExposure[[Europe]:[Other]]").NumberFormat = "0.0%"
combined_sheet.Range("MarketExposure[[Large Cap]:[Small Cap]]").NumberFormat = "0.0%"
combined_sheet.Range("GrossExposure[[Technology Attrib]:[Other Attrib]]").NumberFormat = "0.00%"
combined_sheet.Range("RegionExposure[[Europe Attrib]:[Other Attrib]]").NumberFormat = "0.00%"
combined_sheet.Range("MarketExposure[[Large Cap Attrib]:[Small Cap Attrib]]").NumberFormat = "0.00%"


'Update Graph Tables data

'Delete existing data

With graph_sheet.ListObjects("TotalExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("GrossExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("RegionExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("MarketExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

'Copy all table data over from Combined to Graph

combined_sheet.ListObjects("TotalExposure").Range.Resize(combined_sheet.ListObjects("TotalExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("TotalExposure2")

combined_sheet.ListObjects("GrossExposure").Range.Resize(combined_sheet.ListObjects("GrossExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("GrossExposure2")
    
combined_sheet.ListObjects("RegionExposure").Range.Resize(combined_sheet.ListObjects("RegionExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("RegionExposure2")
    
combined_sheet.ListObjects("MarketExposure").Range.Resize(combined_sheet.ListObjects("MarketExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("MarketExposure2")


'On successful run, delete date field to ensure good data isn't overwriten

input_sheet.Range("E4").Value = ""

End Sub


Sub DeleteDateData()

'Declare variables

Dim wb As Workbook
Dim input_sheet, combined_sheet, graph_sheet As Worksheet
    
Dim input_date As Date

Dim rw As Variant

Dim response As Integer


'Setting workbook and worksheet variables

Set wb = ActiveWorkbook
Set input_sheet = wb.Sheets("Input")
Set combined_sheet = wb.Sheets("Combined Tables")
Set graph_sheet = wb.Sheets("Graph Tables")


'Check if input_date is given

If IsDate(input_sheet.Range("E4").Value) And Not IsEmpty(input_sheet.Range("E4").Value) Then
    input_date = DateTime.DateSerial(Year(input_sheet.Range("E4").Value), Month(input_sheet.Range("E4").Value), 1)
Else
    MsgBox "Please check Input Date value", , "ALERT"
    Exit Sub
End If


'Warning of deletion

response = MsgBox("Are you sure you would like to delete all data for " & Format(input_date, "m/yyyy") & "?", vbYesNo, "WARNING!")

If response = vbNo Then
    MsgBox "Macro Canceled", , "ALERT"
    Exit Sub
End If
    

'Deleting rows from each table

'Total Exposure
For Each rw In combined_sheet.Range("TotalExposure[Date]")
    If rw = input_date Then
        combined_sheet.ListObjects("TotalExposure").ListRows(rw.Row - 1).Delete
        Exit For
    End If
    
    Next rw

rw = CDate(1 / 1 / 1900)

'Gross Exposure
For Each rw In combined_sheet.Range("GrossExposure[Date]")
    If rw = input_date Then
        combined_sheet.ListObjects("GrossExposure").ListRows(rw.Row - 1).Delete
        Exit For
    End If
    
    Next rw

rw = CDate(1 / 1 / 1900)

'Region Exposure

For Each rw In combined_sheet.Range("RegionExposure[Date]")
    If rw = input_date Then
        combined_sheet.ListObjects("RegionExposure").ListRows(rw.Row - 1).Delete
        Exit For
    End If
    
    Next rw

rw = CDate(1 / 1 / 1900)

'Market Exposure

For Each rw In combined_sheet.Range("MarketExposure[Date]")
    If rw = input_date Then
        combined_sheet.ListObjects("MarketExposure").ListRows(rw.Row - 1).Delete
        Exit For
    End If
    
    Next rw

rw = CDate(1 / 1 / 1900)


'Update Graph Tables data

'Delete existing data

With graph_sheet.ListObjects("TotalExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("GrossExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("RegionExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("MarketExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

'Copy all table data over from Combined to Graph

combined_sheet.ListObjects("TotalExposure").Range.Resize(combined_sheet.ListObjects("TotalExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("TotalExposure2")

combined_sheet.ListObjects("GrossExposure").Range.Resize(combined_sheet.ListObjects("GrossExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("GrossExposure2")
    
combined_sheet.ListObjects("RegionExposure").Range.Resize(combined_sheet.ListObjects("RegionExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("RegionExposure2")
    
combined_sheet.ListObjects("MarketExposure").Range.Resize(combined_sheet.ListObjects("MarketExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("MarketExposure2")


'On successful run, delete date field to ensure good data isn't overwriten

input_sheet.Range("E4").Value = ""

End Sub


Sub GraphRangeMacro()

'Declare variables

Dim wb As Workbook
Dim input_sheet, combined_sheet, graph_sheet As Worksheet
    
Dim begin_date, end_date As Date

Dim rw As Variant


'Setting workbook and worksheet variables

Set wb = ActiveWorkbook
Set input_sheet = wb.Sheets("Input")
Set combined_sheet = wb.Sheets("Combined Tables")
Set graph_sheet = wb.Sheets("Graph Tables")


'Update Graph Tables data

'Delete existing data

With graph_sheet.ListObjects("TotalExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("GrossExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("RegionExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

With graph_sheet.ListObjects("MarketExposure2")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

'Copy all table data over from Combined to Graph

combined_sheet.ListObjects("TotalExposure").Range.Resize(combined_sheet.ListObjects("TotalExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("TotalExposure2")

combined_sheet.ListObjects("GrossExposure").Range.Resize(combined_sheet.ListObjects("GrossExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("GrossExposure2")
    
combined_sheet.ListObjects("RegionExposure").Range.Resize(combined_sheet.ListObjects("RegionExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("RegionExposure2")
    
combined_sheet.ListObjects("MarketExposure").Range.Resize(combined_sheet.ListObjects("MarketExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
    Destination:=graph_sheet.ListObjects("MarketExposure2")


'Check if input_date is given

If IsDate(input_sheet.Range("E7").Value) And Not IsEmpty(input_sheet.Range("E7").Value) Then
    begin_date = DateTime.DateSerial(Year(input_sheet.Range("E7").Value), Month(input_sheet.Range("E7").Value), 1)
ElseIf IsEmpty(input_sheet.Range("E7").Value) Then
    begin_date = CDate(1 / 1 / 1900)
Else
    MsgBox "Please check beginning Graph Date Range value", , "ALERT"
    Exit Sub
End If

If IsDate(input_sheet.Range("F7").Value) And Not IsEmpty(input_sheet.Range("F7").Value) Then
    end_date = DateTime.DateSerial(Year(input_sheet.Range("F7").Value), Month(input_sheet.Range("F7").Value), 1)
ElseIf IsEmpty(input_sheet.Range("F7").Value) Then
    end_date = CDate(12 / 12 / 3000)
Else
    MsgBox "Please check ending Graph Date Range value", , "ALERT"
    Exit Sub
End If


'Check if entire date range

If begin_date = CDate(1 / 1 / 1900) And end_date = CDate(12 / 12 / 3000) Then
    'Update Graph Tables data

    'Delete existing data
    
    With graph_sheet.ListObjects("TotalExposure2")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With graph_sheet.ListObjects("GrossExposure2")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With graph_sheet.ListObjects("RegionExposure2")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With graph_sheet.ListObjects("MarketExposure2")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    'Copy all table data over from Combined to Graph
    
    combined_sheet.ListObjects("TotalExposure").Range.Resize(combined_sheet.ListObjects("TotalExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
        Destination:=graph_sheet.ListObjects("TotalExposure2")
    
    combined_sheet.ListObjects("GrossExposure").Range.Resize(combined_sheet.ListObjects("GrossExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
        Destination:=graph_sheet.ListObjects("GrossExposure2")
        
    combined_sheet.ListObjects("RegionExposure").Range.Resize(combined_sheet.ListObjects("RegionExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
        Destination:=graph_sheet.ListObjects("RegionExposure2")
        
    combined_sheet.ListObjects("MarketExposure").Range.Resize(combined_sheet.ListObjects("MarketExposure").DataBodyRange.Rows.Count).Offset(1).Copy _
        Destination:=graph_sheet.ListObjects("MarketExposure2")
    
    Exit Sub
End If


'Deleting rows from each table

'Total Exposure
For Each rw In graph_sheet.Range("TotalExposure2[Date]")
    If Not (rw >= begin_date And rw <= end_date) Then
        graph_sheet.ListObjects("TotalExposure2").ListRows(rw.Row - 1).Delete
    End If

    Next rw

rw = CDate(1 / 1 / 1900)

'Gross Exposure
For Each rw In graph_sheet.Range("GrossExposure2[Date]")
    If Not (rw >= begin_date And rw <= end_date) Then
        graph_sheet.ListObjects("GrossExposure2").ListRows(rw.Row - 1).Delete
    End If

    Next rw

rw = CDate(1 / 1 / 1900)

'Region Exposure

For Each rw In graph_sheet.Range("RegionExposure2[Date]")
    If Not (rw >= begin_date And rw <= end_date) Then
        graph_sheet.ListObjects("RegionExposure2").ListRows(rw.Row - 1).Delete
    End If

    Next rw

rw = CDate(1 / 1 / 1900)

'Market Exposure

For Each rw In graph_sheet.Range("MarketExposure2[Date]")
    If Not (rw >= begin_date And rw <= end_date) Then
        graph_sheet.ListObjects("MarketExposure2").ListRows(rw.Row - 1).Delete
    End If

    Next rw

End Sub
