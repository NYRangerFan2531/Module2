Attribute VB_Name = "Module1"
Sub StockAnalysisMultiYear()

' Rutgers Data Analysis Course
' Module 2 Challange: Stock Analysis
' Code by Leonid Lyakhovich

Dim wb As Workbook
Dim SheetAnalysis As Worksheet
Dim SheetReport As Worksheet
Dim tickerName As String
Dim Volume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim LastDataRow As Double
Dim LastReportRow As Double

Dim largeInc As Double
Dim largeDec As Double
Dim largeVolm As Double

Dim largeIncTic, largeDecTic, largeVolmTic As String


Set wb = ThisWorkbook

For Each SheetAnalysis In wb.Sheets
    
    'Set Up Reporting Sheet
    Sheets.Add.Name = "Report_" & SheetAnalysis.Name
    
    Set SheetReport = Sheets("Report_" & SheetAnalysis.Name)
    SheetReport.Range("A1").Value = "Ticker"
    SheetReport.Range("B1").Value = "Yearly Change"
    SheetReport.Range("C1").Value = "Percent Change"
    SheetReport.Range("D1").Value = "Total Stock Value"
    
    
    'Get Last Row of Data
    
    LastDataRow = SheetAnalysis.Cells(Rows.Count, 1).End(xlUp).Row
    
    Volume = 0 'Set Volume to Zero
    OpenPrice = SheetAnalysis.Cells(2, 3).Value   ' Open Price
    tickerName = SheetAnalysis.Cells(2, 1).Value  ' Get First Stock Ticker
    
    'Get Data for the Tickers and Report Out
    For Row = 2 To LastDataRow:
    
        Volume = Volume + SheetAnalysis.Cells(Row, 7).Value
        
        If tickerName <> SheetAnalysis.Cells(Row + 1, 1).Value Then
            ' if the next ticker is different, end of data reached and report out
            SheetReport.Select
            
            'determine the last row in report
            LastReportRow = SheetReport.Cells(Rows.Count, 1).End(xlUp).Row
            SheetReport.Cells(LastReportRow + 1, 1).Value = tickerName ' ticker Name
            SheetReport.Cells(LastReportRow + 1, 4).Value = Volume ' Volume
            
            ClosePrice = SheetAnalysis.Cells(Row, 6).Value ' Close Price
            
            YearlyChange = ClosePrice - OpenPrice
            
            SheetReport.Cells(LastReportRow + 1, 2).Value = YearlyChange
            SheetReport.Cells(LastReportRow + 1, 3).Value = YearlyChange / OpenPrice
            
            'Formating
            SheetReport.Cells(LastReportRow + 1, 3).NumberFormat = "0.00%"
            
            If YearlyChange > 0 Then
                SheetReport.Cells(LastReportRow + 1, 2).Interior.ColorIndex = 4 ' Green
            Else
                SheetReport.Cells(LastReportRow + 1, 2).Interior.ColorIndex = 3  'Red
            End If
            
            'set up for next ticker
            tickerName = SheetAnalysis.Cells(Row + 1, 1).Value 'Get Next Ticker Name
            Volume = 0 'Set Volume to Zero
            OpenPrice = SheetAnalysis.Cells(Row + 1, 3).Value  ' Get Open Price
        End If
        
    Next Row
    SheetReport.Columns("A:D").AutoFit ' AutoFit Reporting Cells
    
    
    
    
    'Analysis all the tickets
    
    'Assign the First Stock to be the largerst increase, decrease, and largest volume
    largeInc = SheetReport.Range("C2").Value
    largeIncTic = SheetReport.Range("A2").Value
    largeDec = SheetReport.Range("C2").Value
    largeDecTic = SheetReport.Range("A2").Value
    largeVolm = SheetReport.Range("D2").Value
    largeVolmTic = SheetReport.Range("A2").Value
    
    
    For Index = 3 To LastReportRow: ' Analysis all stocks except the first one and title rows
        
        tempChange = SheetReport.Cells(Index, 3).Value
        tempTic = SheetReport.Cells(Index, 1).Value
        tempVol = SheetReport.Cells(Index, 4).Value
        
        If tempChange > largeInc Then
    'Check if stock increase is larger then previous largerst, if yes, it is the new largerst increase
            largeInc = tempChange
            largeIncTic = tempTic
        ElseIf tempChange < largeDec Then
    'Check if stock decrease is smaller (more negative) then previous largerst, if yes, it is the new largerst decrease
            largeDec = tempChange
            largeDecTic = tempTic
        End If
        
        If tempVol > largeVolm Then
    'Check if stock Volume is larger then previous largerst, if yes, it is the new largerst Volume
            largeVolm = tempVol
            largeVolmTic = tempTic
        End If
    Next Index
    
    'Report Ticker Statisitc
    SheetReport.Range("I1") = "Ticker"
    SheetReport.Range("J1") = "Value"
    SheetReport.Range("H2") = "Greatest % Increase"
    SheetReport.Range("I2") = largeIncTic
    SheetReport.Range("J2").NumberFormat = "0.00%"
    SheetReport.Range("J2") = largeInc
    SheetReport.Range("H3") = "Greatest % Decrease"
    SheetReport.Range("I3") = largeDecTic
    SheetReport.Range("J3").NumberFormat = "0.00%"
    SheetReport.Range("J3") = largeDec
    SheetReport.Range("H4") = "Greatest Total Volume"
    SheetReport.Range("I4") = largeVolmTic
    SheetReport.Range("J4") = largeVolm
    
    SheetReport.Columns("H:J").AutoFit ' AutoFit Reporting Cells

Next SheetAnalysis

End Sub

