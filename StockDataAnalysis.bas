Attribute VB_Name = "Module1"
Sub StockData()

'Identify variables in Stock data

    Dim wksht As Worksheet
    
    'Stock data chart from start to finish
    
    Dim Bgintable As Double
    Dim LastRow As Double
    
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim BginDate As Long
    Dim LastDate As Long
    
   'Summary table variables
   
    Dim SummaryRow As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As LongLong
    
    'Greatest increase/decrease variables
    
    Dim GreatestIncreaseV As Double
    Dim GreatestDecreaseV As Double
    Dim GreatestIncreaseT As String
    Dim GreatestDecreaseT As String
    Dim GreatestVolumeV As Double
    Dim GreatestVolumeT As String
    
    ' To loop through all the worksheets
    
    For Each wksht In Worksheets
    
    'Creating Column: Summary Table of the Stock Data
        'Ticker Symbol, Yearly Change, Percent Change, Total Stock Volume
        
    wksht.Range("I1, P1").Value = "Ticker"
    wksht.Cells(1, 10).Value = "Yearly Change"
    wksht.Cells(1, 11).Value = "Percent Change"
    wksht.Cells(1, 12).Value = "Total Stock Volume"
    
    'Ceating Columm: Greatest Increase/Decrease Table
    
    wksht.Cells(1, 17).Value = "Value"
    wksht.Cells(2, 15).Value = "Greatest % Increase"
    wksht.Cells(3, 15).Value = "Greatest % Decrease"
    wksht.Cells(4, 15).Value = "Greatest Total Volume"
    
    'To Nested loop for each worksheet; identify the last row
    
    LastRow = wksht.Cells(Rows.Count, 1).End(xlUp).Row
    
    'We need a starting point to analyze the Stock date
        'So we are going to start at (A,2), which is 2.
        
    Bgintable = 2
    
    
    ' Set Values
    
    BginDate = 99999999
    TotalStockVolume = 0
    GreatestIncreaseV = -99999
    GreatestDecreaseV = 99999
    GreatestVolumeV = -1
    
    ' We have to create a nested loop to go through the data in the worksheet
    
    For SummaryRow = 2 To LastRow
    
    Ticker = wksht.Cells(SummaryRow, 1).Value
    LastDate = wksht.Cells(SummaryRow, 2).Value
    TotalStockVolume = TotalStockVolume + wksht.Cells(SummaryRow, 7).Value
    
    'If Statement
    
    If Ticker = wksht.Cells(SummaryRow + 1, 1).Value Then
    
    If LastDate < BginDate Then
        BginDate = LastDate
        OpenValue = wksht.Cells(SummaryRow, 3).Value
    End If
    
    
    Else
        CloseValue = wksht.Cells(SummaryRow, 6).Value
        YearlyChange = CloseValue - OpenValue
    
    If OpenValue <> 0 Then
        PercentChange = (YearlyChange / OpenValue)
        
    Else
        PercentChange = 0
        
    End If
        
    
    'If Statement
    
    If TotalStockVolume > GreatestVolumeV Then
    GreatestVolumeV = TotalStockVolume
    GreatestVolumeT = Ticker
    End If
    
    If PercentChange > GreatestIncreaseV Then
    GreatestIncreaseV = PercentChange
    GreatestIncreaseT = Ticker
    End If
    
    If PercentChange < GreatestDecreaseV Then
    GreatestDecreaseV = PercentChange
    GreatestDecreaseT = Ticker
    End If
    
    'Populate output values onto Summary Table
    
    wksht.Cells(Bgintable, 9).Value = Ticker
    wksht.Cells(Bgintable, 10).Value = YearlyChange
    
    
    'Format Yearly Change with colors
    
    'If Statement
    If YearlyChange >= 0 Then
    wksht.Cells(Bgintable, 10).Interior.Color = RGB(0, 255, 0) 'Green
    wksht.Cells(Bgintable, 11).Interior.Color = RGB(0, 255, 0) 'Green
    
    Else
    
    wksht.Cells(Bgintable, 10).Interior.Color = RGB(255, 0, 0) 'Red
    wksht.Cells(Bgintable, 11).Interior.Color = RGB(255, 0, 0) 'Red
    
    End If
    
    wksht.Cells(Bgintable, 11).Value = FormatPercent(PercentChange)
    wksht.Cells(Bgintable, 12).Value = TotalStockVolume
    
    'reset the values
    TotalStockVolume = 0
    BginDate = 99999999
    
    'Output to the next table of data start
    Bgintable = Bgintable + 1
    
    End If
    
    Next SummaryRow
    
    
    'Fill in the Greatest Increase/Decrease Table
    
    wksht.Range("P4").Value = GreatestVolumeT
    wksht.Range("Q4").Value = GreatestVolumeV
    wksht.Range("P2").Value = GreatestIncreaseT
    wksht.Range("Q2").Value = FormatPercent(GreatestIncreaseV)
    wksht.Range("P3").Value = GreatestDecreaseT
    wksht.Range("Q3").Value = FormatPercent(GreatestDecreaseV)
    
    wksht.Range("A1:Q1").EntireColumn.AutoFit
    wksht.Range("I1:Q1").Font.Bold = True
    
    Next wksht
    
End Sub
