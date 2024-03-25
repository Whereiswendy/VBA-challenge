Attribute VB_Name = "Module1"
Sub ProcessMultipleSheets()
    ' This subroutine dynamically processes multiple sheets within the workbook,
    ' excluding "Introduction" and "Summary" sheets. It performs stock data analysis and applies conditional formatting.
    ' After processing all relevant sheets, it calls MergeMaxMinEval to merge and evaluate the data.
    
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Set sheetNames = New Collection
    
    ' Dynamically build the list of sheet names to process.
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Introduction" And ws.Name <> "Summary" Then
            sheetNames.Add ws.Name
        End If
    Next ws
    
    ' Process each sheet collected, excluding "Introduction" and "Summary".
    Dim sheetName As Variant
    For Each sheetName In sheetNames
        ' Set the worksheet object to the current sheet in the loop.
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Call subroutines to process stock data and apply conditional formatting on the current sheet.
        LoopThroughStocks ws
        ApplyConditionalFormatting ws
    Next sheetName
    
    ' Call MergeMaxMinEval to perform data merging and evaluation after processing all specified sheets.
    MergeMaxMinEval
End Sub
Sub ResetMultipleSheets()
    ' This subroutine resets data across multiple sheets specified in the array.
    ' It clears specified columns in each sheet using the clearResults subroutine.
    ' Additionally, it checks for the existence of a "Summary" sheet and deletes it if found.
    
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Dim summarySheet As Worksheet
    Set sheetNames = New Collection
    
    ' Dynamically build the list of sheet names to process.
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Introduction" And ws.Name <> "Summary" Then
            sheetNames.Add ws.Name
        End If
    Next ws
    
    ' Process each sheet collected, excluding "Introduction" and "Summary".
    Dim sheetName As Variant
    For Each sheetName In sheetNames
        ' Set the worksheet object to the current sheet in the loop.
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Call clearResults to clear specified columns on the current sheet.
        clearResults ws
    Next sheetName
    
    ' Disable Excel's warning prompts to allow for silent deletion of the "Summary" sheet.
    Application.DisplayAlerts = False
    
    ' Attempt to set the summarySheet variable to the "Summary" worksheet.
    ' If "Summary" doesn't exist, this will error, which is why error handling is temporarily disabled.
    On Error Resume Next
    Set summarySheet = ThisWorkbook.Sheets("Summary")
    
    ' Re-enable normal error handling.
    On Error GoTo 0
    
    ' If the "Summary" sheet was found (meaning summarySheet is not Nothing), then delete it.
    If Not summarySheet Is Nothing Then
        summarySheet.Delete
    End If
    
    ' Re-enable Excel's warning prompts after attempting to delete the "Summary" sheet.
    Application.DisplayAlerts = True
End Sub
Sub clearResults(ws As Worksheet)
    ' This subroutine clears the contents of specified columns and removes any conditional formatting from column J
    ' on the given worksheet (ws). It's typically used to reset the data before running new analyses.
    
    ' Define the columns to clear. You can modify this array to target different columns as needed.
    Dim cols As Variant
    cols = Array("I", "J", "K", "L") ' Columns I, J, K, and L are specified for clearing.
    
    ' Loop through each specified column in the cols array.
    For Each col In cols
        ' Construct the range to clear for the current column, from row 1 to the last row with data.
        ' This uses the .End(xlUp) method to find the last non-empty cell in the column and clears all cells above it.
        ws.Range(col & "1:" & col & ws.Cells(ws.Rows.Count, col).End(xlUp).row).ClearContents
    Next col
    
    ' In addition to clearing the contents, this subroutine also removes any conditional formatting that might be
    ' present in column J. This ensures that old formatting does not affect the appearance of new or updated data.
    ws.Range("J:J").FormatConditions.Delete
End Sub
Sub LoopThroughStocks(ws As Worksheet)
    ' This subroutine processes stock data on a given worksheet. It calculates the yearly change,
    ' percent change from opening to closing price, and total stock volume for each stock. It sorts
    ' the stock data, sets up headers, and applies formatting.

    ' Sort the data before processing.
    SortData ws

    ' Setup headers for the output.
    SetupHeaders ws

    ' Initialize tracking variables.
    Dim lastRow As Long, outputRow As Long, i As Long
    Dim currentStockTicker As String, previousStockTicker As String
    Dim openingPrice As Double, closingPrice As Double
    Dim yearlyChange As Double, percentChange As Double

    ' Determine the last row with data in column A.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Set the initial output row and an unrealistic initial value for the opening price.
    outputRow = 2
    openingPrice = -1
    previousStockTicker = ""

    ' Loop through each row starting from the second to process stock data.
    For i = 2 To lastRow + 1 ' +1 to include logic for the final stock ticker.
        currentStockTicker = ws.Cells(i, "A").Value

        ' Process when a new stock ticker is encountered or on the final iteration.
        If currentStockTicker <> previousStockTicker And i > 2 Or i = lastRow + 1 Then
            ' Finalize calculations for the previous stock ticker.
            If openingPrice <> -1 Then
                ' Ensure closing price is from the correct row, accommodating the last iteration.
                closingPrice = IIf(i = lastRow + 1, ws.Cells(i - 1, "F").Value, ws.Cells(i - 1, "F").Value)
                yearlyChange = closingPrice - openingPrice
                percentChange = IIf(openingPrice <> 0, (yearlyChange / openingPrice) * 100, 0)
                
                ' Output the calculated values.
                ws.Cells(outputRow, "J").Value = yearlyChange
                ws.Cells(outputRow, "K").Value = percentChange
                ws.Cells(outputRow, "L").Value = Application.WorksheetFunction.SumIf(ws.Range("A2:A" & lastRow), previousStockTicker, ws.Range("G2:G" & lastRow))
                
                ' Increment the output row for the next set of data.
                outputRow = outputRow + 1
            End If
            ' Reset opening price for the new stock ticker.
            openingPrice = -1
        End If

        ' Setup for a new stock ticker.
        If currentStockTicker <> previousStockTicker And currentStockTicker <> "" Then
            openingPrice = ws.Cells(i, "C").Value ' Set opening price for the new stock.
            previousStockTicker = currentStockTicker ' Update the tracker.
            ws.Cells(outputRow, "I").Value = currentStockTicker ' Write the ticker to column I.
        End If
    Next i
    
    ' Format percent changes in column K.
    ws.Range("K2:K" & outputRow - 1).NumberFormat = "0.00%"
End Sub
Sub SortData(ws As Worksheet)
    ' Sort stock data based on ticker symbol (Column A) then date (Column B).
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row), Order:=xlAscending
        .SetRange ws.Range("A1:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row)
        .Header = xlYes
        .Apply
    End With
    ' Display a message indicating the completion of the sorting process for the current worksheet.
    MsgBox "Processed data for sheet: " & ws.Name, vbInformation, "Data Processed"
End Sub
Sub SetupHeaders(ws As Worksheet)
    ' Configure and format headers for the output columns.
    Dim headers As Variant
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    Dim col As Integer
    For col = 0 To UBound(headers)
        With ws.Cells(1, col + 9) ' Output starts at column "I".
            .Value = headers(col)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next col
End Sub
Sub ApplyConditionalFormatting(ws As Worksheet)
    ' This subroutine applies conditional formatting to the "Yearly Change" column (J) in the specified worksheet.
    ' Negative changes are highlighted in red, and positive changes are highlighted in green.
    
    ' Determine the last row with data in column J to ensure formatting applies to all relevant cells.
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).row
    
    ' Target the range from J2 to the last row with data in column J.
    With ws.Range("J2:J" & lastRow)
        ' First, clear any existing conditional formats to avoid duplication.
        .FormatConditions.Delete
        
        ' Apply a conditional formatting rule for negative values:
        ' If a cell's value is less than 0, its background is set to red with white text for readability.
        With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Set the cell background to red.
            .Font.Color = RGB(255, 255, 255) ' Set the font color to white.
            .SetFirstPriority ' Ensure this rule takes precedence.
        End With
        
        ' Apply a conditional formatting rule for positive values or zero:
        ' If a cell's value is greater than or equal to 0, its background is set to green with white text.
        With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
            .Interior.Color = RGB(50, 205, 50) ' Set the cell background to green.
            .Font.Color = RGB(255, 255, 255) ' Set the font color to white.
        End With
    End With
End Sub
Sub MergeMaxMinEval()
    ' Creates or resets a Summary sheet, then consolidates data from multiple sheets within the workbook.
    ' Finally, it applies conditional formatting and calculates specific metrics such as max percent increase.

    ' Prepare the Summary sheet using the renamed subroutine.
    Dim summarySheet As Worksheet
    SetupSummarySheet summarySheet
    
    ' Define headers for the Summary sheet and apply formatting.
    SetupSummaryHeaders summarySheet
    
    ' Consolidate data from specified sheets into the Summary sheet.
    ConsolidateDataIntoSummary summarySheet
    
    ' Apply conditional formatting to highlight negative and positive changes.
    ApplySummaryConditionalFormatting summarySheet
    
    ' Calculate and display metrics: Max Percent Increase, Min Percent Decrease, and Max Total Volume.
    DisplayMetrics summarySheet
End Sub
Sub PrepareSummarySheet(ByRef summarySheet As Worksheet)
    ' Handles creation or resetting of the Summary sheet.
    Application.DisplayAlerts = False ' Disable warning prompts for deletion.
    On Error Resume Next ' Suppress errors for attempting to reference a non-existent sheet.
    Set summarySheet = ThisWorkbook.Sheets("Summary")
    If Not summarySheet Is Nothing Then summarySheet.Delete ' Delete if exists.
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Introduction"))
    summarySheet.Name = "Summary"
    Application.DisplayAlerts = True ' Re-enable warning prompts.
    On Error GoTo 0 ' Re-enable regular error handling.
End Sub
Sub SetupSummaryHeaders(summarySheet As Worksheet)
    ' Sets up the headers on the Summary sheet.
    Dim headers As Variant
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With summarySheet.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
End Sub
Sub ConsolidateDataIntoSummary(ByRef summarySheet As Worksheet)
    
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Set sheetNames = New Collection
    
    ' Dynamically build the list of sheet names to process, excluding "Introduction" and "Summary".
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Introduction" And ws.Name <> "Summary" Then
            sheetNames.Add ws.Name
        End If
    Next ws
    
    ' Call the renamed subroutine to setup the Summary sheet.
    SetupSummarySheet summarySheet
    
     ' Process each dynamically collected sheet, excluding "Introduction" and "Summary".
    Dim sheetName As Variant
    For Each sheetName In sheetNames
        ' Set the worksheet object to the current sheet in the loop.
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Consolidate data from the current sheet into the Summary sheet.
        Dim lastRow As Long
        Dim lastRowSummary As Long
        Dim i As Long, j As Long
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        
        For i = 2 To lastRow ' Skip headers, assuming data starts from row 2.
            lastRowSummary = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).row + 1
            For j = 1 To 4 ' Map columns I to L from source to A to D in summary.
                summarySheet.Cells(lastRowSummary, j).Value = ws.Cells(i, j + 8).Value
            Next j
            ' Add the sheet name in column E for each row of consolidated data.
            summarySheet.Cells(lastRowSummary, 5).Value = ws.Name
        Next i
    Next sheetName
End Sub
Sub SetupSummarySheet(ByRef summarySheet As Worksheet)
    ' Checks every sheet in the workbook to see if the "Summary" sheet already exists.
    Dim doesExist As Boolean
    doesExist = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Summary" Then
            doesExist = True
            Exit For
        End If
    Next ws
    
    ' If the "Summary" sheet exists, clear it. If not, create it after the "Introduction" sheet.
    If doesExist Then
        Set summarySheet = ThisWorkbook.Sheets("Summary")
        summarySheet.Cells.ClearContents ' Clears the content but keeps formatting and shapes.
    Else
        Dim introSheet As Worksheet
        Set introSheet = ThisWorkbook.Sheets("Introduction")
        Set summarySheet = ThisWorkbook.Sheets.Add(After:=introSheet)
        summarySheet.Name = "Summary"
    End If
    ' Call the subroutine to initialize headers after setting up or clearing the Summary sheet.
    InitializeSummaryHeaders summarySheet
End Sub
Sub InitializeSummaryHeaders(ByRef summarySheet As Worksheet)
    ' Sets up the headers on the Summary sheet and applies the required formatting.
    
    Dim headers As Variant
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        With summarySheet.Cells(1, i + 1) ' Headers start from column A which is 1 in index
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
End Sub
Sub ApplySummaryConditionalFormatting(summarySheet As Worksheet)
    ' Applies conditional formatting for the "Yearly Change" column in the Summary sheet.
    Dim lastRow2 As Long
    lastRow2 = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).row
    With summarySheet.Range("B2:B" & lastRow2)
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0").Interior.Color = RGB(255, 0, 0)
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0").Interior.Color = RGB(50, 205, 50)
    End With
End Sub
Sub DisplayMetrics(summarySheet As Worksheet)
    ' Calculates and displays metrics for the greatest percent increase, decrease, and total volume.
    ' Includes formatting adjustments for readability.
    Dim maxPercentIncrease As Double, minPercentDecrease As Double, maxTotalVolume As Double
    Dim tickerMaxIncrease As String, tickerMinDecrease As String, tickerMaxVolume As String
    
    ' Calculations for max/min percent change and total volume.
    maxPercentIncrease = Application.WorksheetFunction.Max(summarySheet.Range("C:C"))
    minPercentDecrease = Application.WorksheetFunction.Min(summarySheet.Range("C:C"))
    maxTotalVolume = Application.WorksheetFunction.Max(summarySheet.Range("D:D"))
    
    ' Identifying tickers associated with the calculated metrics.
    IdentifyTickers summarySheet, maxPercentIncrease, minPercentDecrease, maxTotalVolume, _
                    tickerMaxIncrease, tickerMinDecrease, tickerMaxVolume
    
    ' Populate the Summary sheet with calculated values and metrics.
    With summarySheet
        .Cells(2, "G").Value = tickerMaxIncrease
        .Cells(3, "G").Value = tickerMinDecrease
        .Cells(4, "G").Value = tickerMaxVolume
        
        .Cells(2, "H").Value = maxPercentIncrease
        .Cells(3, "H").Value = minPercentDecrease
        .Cells(4, "H").Value = maxTotalVolume

        ' Formatting for metrics display.
        .Range("G1:H1").Value = Array("Ticker", "Value")
        .Range("F2:F4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        .Range("F1:H1").Font.Bold = True
        .Range("F2:F4").Font.Bold = True
        .Range("F1:H4").HorizontalAlignment = xlCenter
        .Range("F1:H4").VerticalAlignment = xlCenter
        .Range("F1:H4").WrapText = True
        .Range("H2:H3").NumberFormat = "0.00"
        .Columns("H").AutoFit
    End With

    ' Additional formatting to ensure numeric values are displayed correctly and clearly.
    EnsureNumericDisplay summarySheet
End Sub
Sub IdentifyTickers(summarySheet As Worksheet, ByVal maxPercentIncrease As Double, ByVal minPercentDecrease As Double, _
                    ByVal maxTotalVolume As Double, ByRef tickerMaxIncrease As String, _
                    ByRef tickerMinDecrease As String, ByRef tickerMaxVolume As String)
    ' Identifies tickers associated with the max percent increase, min percent decrease, and max total volume.
    Dim row As Long
    For row = 2 To summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).row
        If summarySheet.Cells(row, "C").Value = maxPercentIncrease Then
            tickerMaxIncrease = summarySheet.Cells(row, "A").Value
        ElseIf summarySheet.Cells(row, "C").Value = minPercentDecrease Then
            tickerMinDecrease = summarySheet.Cells(row, "A").Value
        End If
        If summarySheet.Cells(row, "D").Value = maxTotalVolume Then
            tickerMaxVolume = summarySheet.Cells(row, "A").Value
        End If
    Next row
End Sub
Sub EnsureNumericDisplay(summarySheet As Worksheet)
    ' Ensures that numeric values, especially those in column H, are displayed in a non-exponential format
    ' if they are large numbers. Adjusts column width as necessary.
    With summarySheet
        ' Adjust column width for the "Value" column to accommodate large numbers without resorting to scientific notation.
        .Columns("H").AutoFit
        
        ' Check if any adjustments are needed for the display format of large numbers.
        .Cells(4, "H").NumberFormat = "#,##0"  ' This format adds thousand separators without decimals for the total volume.
    End With
End Sub




