VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateQuarterlyStockDataAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim currentQuarter As String
    Dim outputRow As Long
    
    ' Loop through each quarter
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        outputRow = 2
        
        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        startRow = 2
        
        Do While startRow <= lastRow
            ticker = ws.Cells(startRow, 1).Value
            currentQuarter = Format(ws.Cells(startRow, 2).Value, "yyyy\Qq")
            
            ' Find the end row for the quarter
            endRow = startRow
            Do While endRow <= lastRow And ws.Cells(endRow, 1).Value = ticker And Format(ws.Cells(endRow, 2).Value, "yyyy\Qq") = currentQuarter
                endRow = endRow + 1
            Loop
            endRow = endRow - 1
            
            ' Calculate quarterly change, percentage change, and total volume
            quarterlyChange = ws.Cells(endRow, 6).Value - ws.Cells(startRow, 3).Value
            percentageChange = (ws.Cells(endRow, 6).Value - ws.Cells(startRow, 3).Value) / ws.Cells(startRow, 3).Value * 100
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).Value = totalVolume
            
            
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            
            outputRow = outputRow + 1
            startRow = endRow + 1
        Loop
    Next ws
End Sub
